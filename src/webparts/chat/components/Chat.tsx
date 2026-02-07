import * as React from 'react';
import { IChatProps, IChatMessage, IChatAttachment, IProject, IStoredChat } from './IChatProps';
import styles from './Chat.module.scss';
import {
  getProjects,
  saveProjects,
  getAllChats,
  saveAllChats,
  getChatsForProject,
  getCurrentProjectId,
  setCurrentProjectId as persistProjectId,
  getCurrentChatId,
  setCurrentChatId as persistChatId,
  messageToStored,
  storedToMessage
} from './ChatStorage';
import {
  getProjectsFromLibrary,
  saveProjectsToLibrary,
  getChatsFromLibrary,
  saveChatToLibrary,
  deleteChatFromLibrary,
  type IChatStorageSharePointConfig
} from './ChatStorageSharePoint';
import * as mammoth from 'mammoth';
import { getDocument, GlobalWorkerOptions } from 'pdfjs-dist/legacy/build/pdf';
import * as XLSX from 'xlsx';

// PDF.js requires workerSrc in browser. Use CDN so we don't bundle the worker (SPFx/browser).
const PDFJS_VERSION = '2.16.105';
if (typeof GlobalWorkerOptions !== 'undefined' && !GlobalWorkerOptions.workerSrc) {
  GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${PDFJS_VERSION}/legacy/build/pdf.worker.min.js`;
}

const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';
const DEFAULT_MODEL = 'gpt-4o-mini';
const MODEL_OPTIONS: { value: string; label: string }[] = [
  { value: 'gpt-4o-mini', label: 'GPT-4o mini' },
  { value: 'gpt-4o', label: 'GPT-4o' },
  { value: 'gpt-4o-turbo', label: 'GPT-4o turbo' },
  { value: 'gpt-4-turbo', label: 'GPT-4 turbo' },
  { value: 'gpt-3.5-turbo', label: 'GPT-3.5 turbo' }
];
const MODEL_FOR_SIMPLE = 'gpt-4o-mini';
const MODEL_FOR_COMPLEX = 'gpt-4o';
const MAX_FILES_PER_MESSAGE = 10;
const MAX_IMAGE_SIZE_MB = 4;
const MAX_FILE_SIZE_MB = 10;
const MAX_TOTAL_TEXT_CHARS = 80000;

function generateId(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result as string);
    r.onerror = () => reject(new Error('Failed to read file'));
    r.readAsDataURL(file);
  });
}

function fileToText(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve((r.result as string) || '');
    r.onerror = () => reject(new Error('Failed to read file'));
    r.readAsText(file, 'UTF-8');
  });
}

function isImageFile(file: File): boolean {
  return file.type.startsWith('image/');
}

/** Plain text and other formats we can read as UTF-8 text (no special parser) */
function isLikelyTextFile(file: File): boolean {
  const t = file.type.toLowerCase();
  const name = file.name.toLowerCase();
  if (t.startsWith('text/') || t === 'application/json' || t === 'application/javascript' || t === 'application/xml' || t === 'application/rtf') return true;
  const ext = name.slice(name.lastIndexOf('.'));
  const textExtensions = [
    '.txt', '.md', '.csv', '.tsv', '.json', '.html', '.htm', '.xml', '.js', '.ts', '.tsx', '.jsx',
    '.log', '.yml', '.yaml', '.ini', '.cfg', '.conf', '.env', '.properties', '.sql', '.rtf',
    '.csv', '.tex', '.rst', '.asciidoc', '.adoc'
  ];
  return textExtensions.indexOf(ext) >= 0;
}

function isPdfFile(file: File): boolean {
  return file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');
}

function isDocxFile(file: File): boolean {
  const t = file.type.toLowerCase();
  const name = file.name.toLowerCase();
  return t === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || name.endsWith('.docx');
}

function isXlsxFile(file: File): boolean {
  const t = file.type.toLowerCase();
  const name = file.name.toLowerCase();
  return (
    t === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
    name.endsWith('.xlsx') ||
    name.endsWith('.xls')
  );
}

type FileThumbType = 'pdf' | 'word' | 'excel' | 'text' | 'generic';

function getFileThumbType(file: File): FileThumbType {
  if (isPdfFile(file)) return 'pdf';
  if (isDocxFile(file)) return 'word';
  if (isXlsxFile(file)) return 'excel';
  if (isLikelyTextFile(file)) return 'text';
  return 'generic';
}

/** 48×48 file-type icon (no trademarked logos; suggestive colors only) */
function FileTypeIcon(props: { type: FileThumbType; className?: string }): React.ReactElement {
  const { type, className } = props;
  const viewBox = '0 0 48 48';
  const common = { width: 48, height: 48, viewBox, className, fill: 'none' };
  switch (type) {
    case 'pdf':
      return (
        <svg {...common} xmlns="http://www.w3.org/2000/svg">
          <path d="M10 6h18l10 10v26H10V6z" fill="#E74C3C" />
          <path d="M28 6v10h10" fill="#C0392B" />
          <text x="24" y="32" textAnchor="middle" fill="white" fontSize="10" fontWeight="bold" fontFamily="sans-serif">PDF</text>
        </svg>
      );
    case 'word':
      return (
        <svg {...common} xmlns="http://www.w3.org/2000/svg">
          <path d="M10 6h18l10 10v26H10V6z" fill="#2B579A" />
          <path d="M28 6v10h10" fill="#1E3A5F" />
          <path d="M14 22h20M14 28h14M14 34h20" stroke="white" strokeWidth="1.5" strokeLinecap="round" opacity={0.9} />
        </svg>
      );
    case 'excel':
      return (
        <svg {...common} xmlns="http://www.w3.org/2000/svg">
          <path d="M10 6h18l10 10v26H10V6z" fill="#217346" />
          <path d="M28 6v10h10" fill="#1B5C38" />
          <path d="M14 20v16M22 20v16M30 20v16M38 20v16M14 20h24M14 28h24M14 36h24" stroke="white" strokeWidth="1.2" strokeLinecap="round" opacity={0.9} />
        </svg>
      );
    case 'text':
      return (
        <svg {...common} xmlns="http://www.w3.org/2000/svg">
          <path d="M10 6h28v36H10V6z" fill="#5C6BC0" stroke="#3F51B5" strokeWidth="1" />
          <path d="M14 18h20M14 24h16M14 30h20" stroke="white" strokeWidth="1.2" strokeLinecap="round" opacity={0.9} />
        </svg>
      );
    default:
      return (
        <svg {...common} xmlns="http://www.w3.org/2000/svg">
          <path d="M10 6h18l10 10v26H10V6z" fill="#78909C" />
          <path d="M28 6v10h10" fill="#607D8B" />
          <path d="M14 24h20M14 30h14" stroke="white" strokeWidth="1.2" strokeLinecap="round" opacity={0.8} />
        </svg>
      );
  }
}

/**
 * Heuristic suggestion: simple (use mini) vs complex (use stronger model).
 * No API call — based on length and keywords only.
 */
function suggestModelFromContent(text: string, hasAttachments: boolean): 'simple' | 'complex' {
  const t = (text || '').trim().toLowerCase();
  if (t.length > 400) return 'complex';
  if (hasAttachments) return 'complex'; // documents/images often need deeper analysis
  const complexKeywords = [
    'step by step', 'explain in detail', 'compare and contrast', 'analyze', 'write code',
    'implement', 'debug', 'function that', 'essay', 'outline', 'compare', 'contrast',
    'pros and cons', 'critically', 'reasoning', 'proof', 'derive', 'algorithm'
  ];
  for (const kw of complexKeywords) {
    if (t.indexOf(kw) >= 0) return 'complex';
  }
  return 'simple';
}

function getSuggestedModelValue(text: string, hasAttachments: boolean): string {
  return suggestModelFromContent(text, hasAttachments) === 'complex' ? MODEL_FOR_COMPLEX : MODEL_FOR_SIMPLE;
}

/** Extract text from a PDF file using PDF.js (bundled so it works in SharePoint) */
async function pdfToText(file: File): Promise<string> {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await getDocument({ data: arrayBuffer }).promise;
  const numPages = pdf.numPages;
  const parts: string[] = [];
  for (let p = 1; p <= numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const pageText = (content.items as { str?: string }[]).map((item) => item.str || '').join(' ');
    parts.push(pageText);
  }
  return parts.join('\n\n');
}

/** Extract text from a Word (.docx) file using mammoth */
async function docxToText(file: File): Promise<string> {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer });
  return result.value || '';
}

/** Extract text from an Excel (.xlsx, .xls) file using SheetJS */
async function xlsxToText(file: File): Promise<string> {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const parts: string[] = [];
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const csv = XLSX.utils.sheet_to_csv(sheet);
    if (csv.trim()) {
      parts.push(`Sheet: ${sheetName}\n${csv}`);
    }
  }
  return parts.join('\n\n');
}

/** Simple markdown: **bold**, `code`, ```code block``` */
function renderMarkdown(text: string): React.ReactNode {
  if (!text) return null;
  const parts: React.ReactNode[] = [];
  let key = 0;
  // Code blocks first (```...```)
  const blockRegex = /```(\w*)\n?([\s\S]*?)```/g;
  let lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = blockRegex.exec(text)) !== null) {
    if (m.index > lastIndex) {
      parts.push(renderInlineMarkdown(text.slice(lastIndex, m.index), key));
      key += 1;
    }
    parts.push(<pre key={key} className={styles.codeBlock}><code>{m[2].trim()}</code></pre>);
    key += 1;
    lastIndex = m.index + m[0].length;
  }
  if (lastIndex < text.length) {
    parts.push(renderInlineMarkdown(text.slice(lastIndex), key));
  }
  return parts.length === 1 ? parts[0] : <>{parts}</>;
}

function renderInlineMarkdown(str: string, keyBase: number): React.ReactNode {
  const parts: React.ReactNode[] = [];
  const re = /(\*\*[^*]+\*\*|`[^`]+`)/g;
  let lastIndex = 0;
  let k = keyBase;
  let m: RegExpExecArray | null;
  while ((m = re.exec(str)) !== null) {
    if (m.index > lastIndex) parts.push(<span key={k}>{str.slice(lastIndex, m.index)}</span>);
    k += 1;
    const raw = m[1];
    if (raw.startsWith('**')) {
      parts.push(<strong key={k}>{raw.slice(2, -2)}</strong>);
    } else if (raw.startsWith('`')) {
      parts.push(<code key={k} className={styles.inlineCode}>{raw.slice(1, -1)}</code>);
    }
    k += 1;
    lastIndex = m.index + m[0].length;
  }
  if (lastIndex < str.length) parts.push(<span key={k}>{str.slice(lastIndex)}</span>);
  return parts.length === 0 ? str : parts.length === 1 ? parts[0] : <>{parts}</>;
}

const Chat: React.FC<IChatProps> = (props) => {
  const { apiKey, docsLibraryName, context } = props;
  const useSharePoint = Boolean(docsLibraryName?.trim() && context?.pageContext?.web?.absoluteUrl && context?.spHttpClient);
  const spConfig = React.useMemo<IChatStorageSharePointConfig | null>(() => {
    if (!useSharePoint || !context) return null;
    return {
      spHttpClient: context.spHttpClient,
      webAbsoluteUrl: context.pageContext.web.absoluteUrl,
      libraryTitle: docsLibraryName!.trim()
    };
  }, [useSharePoint, context?.pageContext?.web?.absoluteUrl, docsLibraryName?.trim()]);

  const [messages, setMessages] = React.useState<IChatMessage[]>([]);
  const [input, setInput] = React.useState('');
  const [attachments, setAttachments] = React.useState<IChatAttachment[]>([]);
  const [model, setModel] = React.useState<string>(DEFAULT_MODEL);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [projects, setProjects] = React.useState<IProject[]>(() => (useSharePoint ? [] : getProjects()));
  const [currentProjectId, setCurrentProjectIdState] = React.useState<string | null>(() => getCurrentProjectId());
  const [allChats, setAllChats] = React.useState<IStoredChat[]>([]);
  const [chats, setChats] = React.useState<IStoredChat[]>([]);
  const [currentChatId, setCurrentChatIdState] = React.useState<string | null>(() => getCurrentChatId());
  const [sidebarOpen, setSidebarOpen] = React.useState(() =>
    typeof window !== 'undefined' && window.matchMedia('(max-width: 768px)').matches ? false : true
  );
  const [sidebarMenuOpen, setSidebarMenuOpen] = React.useState<string | null>(null);
  const [dropdownAnchor, setDropdownAnchor] = React.useState<{ top: number; left: number; bottom: number } | null>(null);
  const [storageLoading, setStorageLoading] = React.useState(useSharePoint);
  const [isMobile, setIsMobile] = React.useState(() =>
    typeof window !== 'undefined' && window.matchMedia('(max-width: 768px)').matches
  );
  const messagesEndRef = React.useRef<HTMLDivElement>(null);
  const dropdownRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const mq = window.matchMedia('(max-width: 768px)');
    const update = (): void => {
      setIsMobile(mq.matches);
      if (mq.matches) setSidebarOpen((o) => (o ? false : o));
    };
    mq.addEventListener('change', update);
    return () => mq.removeEventListener('change', update);
  }, []);

  const setCurrentProjectId = React.useCallback((id: string | null) => {
    setCurrentProjectIdState(id);
    persistProjectId(id);
    setCurrentChatIdState(null);
    persistChatId(null);
  }, []);
  const setCurrentChatId = React.useCallback((id: string | null) => {
    setCurrentChatIdState(id);
    persistChatId(id);
  }, []);

  React.useEffect(() => {
    if (!useSharePoint || !spConfig) return;
    let cancelled = false;
    (async () => {
      try {
        const [proj, chatsList] = await Promise.all([
          getProjectsFromLibrary(spConfig),
          getChatsFromLibrary(spConfig)
        ]);
        if (cancelled) return;
        if (proj.length === 0) {
          const defaultProject: IProject = { id: generateId(), name: 'Default', createdAt: new Date().toISOString() };
          await saveProjectsToLibrary(spConfig, [defaultProject]);
          if (cancelled) return;
          setProjects([defaultProject]);
          setCurrentProjectId(defaultProject.id);
          setCurrentProjectIdState(defaultProject.id);
        } else {
          setProjects(proj);
        }
        setAllChats(chatsList);
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : 'Failed to load chats from library');
      } finally {
        if (!cancelled) setStorageLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [useSharePoint, spConfig?.libraryTitle]);

  React.useEffect(() => {
    if (!useSharePoint) {
      let proj = getProjects();
      if (proj.length === 0) {
        const defaultProject: IProject = { id: generateId(), name: 'Default', createdAt: new Date().toISOString() };
        saveProjects([defaultProject]);
        setCurrentProjectId(defaultProject.id);
        setProjects([defaultProject]);
        setCurrentProjectIdState(defaultProject.id);
      }
    }
  }, [useSharePoint]);

  React.useEffect(() => {
    if (!currentProjectId) {
      setChats([]);
      return;
    }
    if (useSharePoint) {
      setChats(allChats.filter((c) => c.projectId === currentProjectId));
    } else {
      setChats(getChatsForProject(currentProjectId));
    }
  }, [currentProjectId, useSharePoint, allChats]);

  React.useEffect(() => {
    if (!currentChatId) {
      setMessages([]);
      return;
    }
    const all = useSharePoint ? allChats : getAllChats();
    const chat = all.find((c) => c.id === currentChatId);
    if (chat) {
      // First chat of a project: we just created it with messages: []. Don't overwrite current state (user message + placeholder).
      if (chat.messages.length > 0) {
        setMessages(chat.messages.map(storedToMessage));
      }
    }
    // When chat not found: do not clear (e.g. SharePoint: not in allChats yet). When found but 0 messages: newly created, don't wipe.
  }, [currentChatId, useSharePoint, allChats]);

  const persistCurrentChat = React.useCallback((msgs: IChatMessage[], title?: string) => {
    const cid = currentChatId;
    const pid = currentProjectId;
    if (!cid || !pid) return;
    const now = new Date().toISOString();
    if (useSharePoint && spConfig) {
      const existing = allChats.find((c) => c.id === cid);
      const titleToUse = title ?? existing?.title ?? 'New chat';
      const updated: IStoredChat = {
        id: cid,
        projectId: pid,
        title: titleToUse,
        messages: msgs.map(messageToStored),
        createdAt: existing?.createdAt ?? now,
        updatedAt: now
      };
      saveChatToLibrary(spConfig, updated).then(() => {
        setAllChats((prev) => {
          const i = prev.findIndex((c) => c.id === cid);
          if (i >= 0) {
            const next = [...prev];
            next[i] = updated;
            return next;
          }
          return [...prev, updated];
        });
      }).catch((e) => setError(e instanceof Error ? e.message : 'Failed to save chat'));
    } else {
      const all = getAllChats();
      const idx = all.findIndex((c) => c.id === cid);
      const titleToUse = title ?? (all[idx]?.title ?? 'New chat');
      const updated: IStoredChat = {
        id: cid,
        projectId: pid,
        title: titleToUse,
        messages: msgs.map(messageToStored),
        createdAt: idx >= 0 ? all[idx].createdAt : now,
        updatedAt: now
      };
      if (idx >= 0) all[idx] = updated;
      else all.push(updated);
      saveAllChats(all);
      setChats(getChatsForProject(pid));
    }
  }, [currentChatId, currentProjectId, useSharePoint, spConfig, allChats]);

  const prevLoadingForPersist = React.useRef(loading);
  React.useEffect(() => {
    if (prevLoadingForPersist.current && !loading && currentChatId && currentProjectId) {
      persistCurrentChat(messages);
      const first = messages.find((m) => m.role === 'user');
      const newTitle = first ? first.content.slice(0, 50).trim() || 'New chat' : null;
      if (useSharePoint && spConfig && newTitle) {
        const c = allChats.find((x) => x.id === currentChatId);
        if (c && c.title === 'New chat') {
          const updated: IStoredChat = { ...c, title: newTitle, updatedAt: new Date().toISOString(), messages: messages.map(messageToStored) };
          saveChatToLibrary(spConfig, updated).then(() => {
            setAllChats((prev) => prev.map((x) => (x.id === currentChatId ? updated : x)));
          }).catch(() => {});
        }
      } else if (!useSharePoint) {
        const all = getAllChats();
        const c = all.find((x) => x.id === currentChatId);
        if (c && c.title === 'New chat' && newTitle) {
          c.title = newTitle;
          c.updatedAt = new Date().toISOString();
          saveAllChats(all);
        }
        setChats(getChatsForProject(currentProjectId));
      }
    }
    prevLoadingForPersist.current = loading;
  }, [loading, currentChatId, currentProjectId, messages, persistCurrentChat, useSharePoint, spConfig, allChats]);

  const textareaRef = React.useRef<HTMLTextAreaElement>(null);
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const abortRef = React.useRef<AbortController | null>(null);

  const scrollToBottom = React.useCallback((smooth = true) => {
    messagesEndRef.current?.scrollIntoView({ behavior: smooth ? 'smooth' : 'auto', block: 'end' });
  }, []);

  React.useEffect(() => {
    scrollToBottom();
  }, [messages, scrollToBottom]);

  const prevLoadingRef = React.useRef(loading);
  React.useEffect(() => {
    const wasLoading = prevLoadingRef.current;
    prevLoadingRef.current = loading;
    if (wasLoading && !loading) {
      const id = requestAnimationFrame(() => {
        requestAnimationFrame(() => scrollToBottom(false));
      });
      return () => cancelAnimationFrame(id);
    }
  }, [loading, scrollToBottom]);

  const newChat = React.useCallback(() => {
    abortRef.current?.abort();
    setInput('');
    setAttachments([]);
    setError(null);
    setLoading(false);
    if (!currentProjectId) return;
    const newId = generateId();
    const newChatRecord: IStoredChat = {
      id: newId,
      projectId: currentProjectId,
      title: 'New chat',
      messages: [],
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
    if (useSharePoint && spConfig) {
      saveChatToLibrary(spConfig, newChatRecord).then(() => {
        setAllChats((prev) => [...prev, newChatRecord]);
      }).catch((e) => setError(e instanceof Error ? e.message : 'Failed to create chat'));
      setCurrentChatId(newId);
      setCurrentChatIdState(newId);
      persistChatId(newId);
      setMessages([]);
    } else {
      const all = getAllChats();
      all.push(newChatRecord);
      saveAllChats(all);
      setChats(getChatsForProject(currentProjectId));
      setCurrentChatId(newId);
      setCurrentChatIdState(newId);
      persistChatId(newId);
      setMessages([]);
    }
  }, [currentProjectId, useSharePoint, spConfig]);

  const removeAttachment = React.useCallback((index: number) => {
    setAttachments((prev) => prev.filter((_, i) => i !== index));
  }, []);

  const onFileChange = React.useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    const list: IChatAttachment[] = [];
    const maxSizeImage = MAX_IMAGE_SIZE_MB * 1024 * 1024;
    const maxSizeFile = MAX_FILE_SIZE_MB * 1024 * 1024;
    let totalChars = 0;
    try {
      for (let i = 0; i < files.length && list.length < MAX_FILES_PER_MESSAGE; i++) {
        const file = files[i];
        const isImage = isImageFile(file);
        const sizeLimit = isImage ? maxSizeImage : maxSizeFile;
        if (file.size > sizeLimit) continue;
        try {
          if (isImage) {
            const dataUrl = await fileToDataUrl(file);
            list.push({ file, dataUrl });
          } else if (isPdfFile(file) && totalChars < MAX_TOTAL_TEXT_CHARS) {
            const text = await pdfToText(file);
            const truncated = text.length > MAX_TOTAL_TEXT_CHARS - totalChars
              ? text.slice(0, MAX_TOTAL_TEXT_CHARS - totalChars) + '\n...[truncated]'
              : text;
            totalChars += truncated.length;
            list.push({ file, textContent: truncated });
          } else if (isDocxFile(file) && totalChars < MAX_TOTAL_TEXT_CHARS) {
            const text = await docxToText(file);
            const truncated = text.length > MAX_TOTAL_TEXT_CHARS - totalChars
              ? text.slice(0, MAX_TOTAL_TEXT_CHARS - totalChars) + '\n...[truncated]'
              : text;
            totalChars += truncated.length;
            list.push({ file, textContent: truncated });
          } else if (isXlsxFile(file) && totalChars < MAX_TOTAL_TEXT_CHARS) {
            const text = await xlsxToText(file);
            const truncated = text.length > MAX_TOTAL_TEXT_CHARS - totalChars
              ? text.slice(0, MAX_TOTAL_TEXT_CHARS - totalChars) + '\n...[truncated]'
              : text;
            totalChars += truncated.length;
            list.push({ file, textContent: truncated });
          } else if (isLikelyTextFile(file) && totalChars < MAX_TOTAL_TEXT_CHARS) {
            const text = await fileToText(file);
            const truncated = text.length > MAX_TOTAL_TEXT_CHARS - totalChars
              ? text.slice(0, MAX_TOTAL_TEXT_CHARS - totalChars) + '\n...[truncated]'
              : text;
            totalChars += truncated.length;
            list.push({ file, textContent: truncated });
          } else {
            list.push({ file });
          }
        } catch (err) {
          const msg = err instanceof Error ? err.message : 'Could not extract text';
          list.push({ file, extractionError: msg });
        }
      }
    } finally {
      if (list.length > 0) {
        setAttachments((prev) => [...prev, ...list].slice(0, MAX_FILES_PER_MESSAGE));
        setError(null);
      } else if (files.length > 0) {
        setError('No files were added. Check size limits (images ≤4 MB, other files ≤10 MB) or try different files.');
      }
      e.target.value = '';
    }
  }, []);

  const buildMessageText = (
    userText: string,
    fileTexts: { name: string; text: string }[],
    otherFileNames?: string[]
  ): string => {
    let text = userText || '';
    if (fileTexts && fileTexts.length > 0) {
      const fileBlocks = fileTexts.map((f) => `--- ${f.name} ---\n${f.text}`).join('\n\n');
      text = text ? `${text}\n\nAttached files:\n\n${fileBlocks}` : `Attached files:\n\n${fileBlocks}`;
    }
    if (otherFileNames && otherFileNames.length > 0) {
      text = (text ? text + '\n\n' : '') + `(Also attached; content not extracted: ${otherFileNames.join(', ')})`;
    }
    return text || '(No content)';
  };

  const buildOpenAIMessages = React.useCallback(
    (
      userText: string,
      imageDataUrls: string[],
      fileTexts: { name: string; text: string }[],
      otherFileNames: string[],
      opts?: { forRegenerate?: boolean }
    ): Array<{ role: string; content: string | object[] }> => {
      const list = opts?.forRegenerate ? messages.slice(0, -1) : messages;
      const apiMessages = list.map((m) => {
        if (m.role === 'assistant') return { role: 'assistant', content: m.content };
        const fullText = buildMessageText(m.content, m.fileTexts || [], m.otherFileNames);
        const parts: object[] = [{ type: 'text', text: fullText }];
        if (m.imageDataUrls?.length) {
          m.imageDataUrls.forEach((url) => parts.push({ type: 'image_url', image_url: { url } }));
        }
        return { role: 'user', content: parts.length === 1 ? fullText : parts };
      });

      if (!opts?.forRegenerate) {
        const newText = buildMessageText(userText, fileTexts, otherFileNames);
        const newParts: object[] = [{ type: 'text', text: newText }];
        imageDataUrls.forEach((url) => newParts.push({ type: 'image_url', image_url: { url } }));
        apiMessages.push({
          role: 'user',
          content: newParts.length === 1 ? newText : newParts
        });
      }
      // When the conversation includes attached file content, tell the model it has access
      const hasAttachedFileContent = apiMessages.some((m) => {
        if (m.role !== 'user') return false;
        const text = typeof m.content === 'string' ? m.content : (Array.isArray(m.content) ? (m.content as { type?: string; text?: string }[]).find((p) => p.type === 'text')?.text : '');
        return typeof text === 'string' && text.includes('Attached files:');
      });
      if (hasAttachedFileContent) {
        apiMessages.unshift({
          role: 'system',
          content: 'The user can attach files. When they do, the full text content of supported files (PDF, Word, Excel, text) is included in the user message under "Attached files:" with the format "--- filename ---" followed by the content. Use that content to answer. Do not say you cannot view or access the files when this content is present.'
        });
      }
      const hasUnsupportedAttachments = !hasAttachedFileContent && apiMessages.some((m) => {
        if (m.role !== 'user') return false;
        const text = typeof m.content === 'string' ? m.content : (Array.isArray(m.content) ? (m.content as { type?: string; text?: string }[]).find((p) => p.type === 'text')?.text : '');
        return typeof text === 'string' && text.includes('content not extracted');
      });
      if (hasUnsupportedAttachments) {
        apiMessages.unshift({
          role: 'system',
          content: 'The user attached files but text extraction failed (e.g. PDF in browser), so you only see "(Also attached; content not extracted: ...)". Do not say you cannot access attachments. Instead, ask the user to paste the relevant text into the chat or to try a different format (e.g. .txt), or suggest using a backend that can extract PDFs.'
        });
      }
      const hasImages = apiMessages.some((m) => {
        if (m.role !== 'user') return false;
        if (!Array.isArray(m.content)) return false;
        return (m.content as { type?: string }[]).some((p) => p.type === 'image_url');
      });
      if (hasImages) {
        apiMessages.unshift({
          role: 'system',
          content: 'The user can attach images. When they do, the images are included in the user message and you can see them. Describe, analyze, or answer questions about what you see in the images. Do not say you are unable to view or recognize images when they have been shared.'
        });
      }
      return apiMessages;
    },
    [messages]
  );

  const sendMessage = React.useCallback(
    async (options?: { regenerate?: boolean }) => {
      const key = apiKey?.trim();
      if (!key) {
        setError('OpenAI API key is not set. Add it in the web part properties (pencil icon).');
        return;
      }

      let userText = input.trim();
      let imageDataUrls = attachments.filter((a) => a.dataUrl).map((a) => a.dataUrl as string);
      let fileTexts = attachments
        .filter((a) => a.textContent !== undefined)
        .map((a) => ({ name: a.file.name, text: a.textContent as string }));
      let otherFileNames = attachments
        .filter((a) => !a.dataUrl && a.textContent === undefined)
        .map((a) => a.file.name);

      if (options?.regenerate && messages.length >= 2) {
        const lastUser = [...messages].reverse().find((m) => m.role === 'user');
        if (lastUser) {
          userText = lastUser.content;
          imageDataUrls = lastUser.imageDataUrls || [];
          fileTexts = lastUser.fileTexts || [];
          otherFileNames = lastUser.otherFileNames || [];
          setMessages((prev) => prev.slice(0, -2));
        }
      }

      const hasContent = userText.length > 0 || imageDataUrls.length > 0 || fileTexts.length > 0 || otherFileNames.length > 0;
      if (!hasContent && !options?.regenerate) return;
      if (loading) return;

      setError(null);
      if (!options?.regenerate) {
        setInput('');
        setAttachments([]);
        const userMessage: IChatMessage = {
          id: generateId(),
          role: 'user',
          content: userText || (imageDataUrls.length > 0 ? '(Sent images)' : '(Sent files)'),
          timestamp: new Date(),
          imageDataUrls: imageDataUrls.length ? imageDataUrls : undefined,
          fileTexts: fileTexts.length ? fileTexts : undefined,
          otherFileNames: otherFileNames.length ? otherFileNames : undefined
        };
        let chatIdToUse = currentChatId;
        if (!currentChatId && currentProjectId) {
          const newId = generateId();
          const title = (userText || '').trim().slice(0, 50) || (imageDataUrls.length ? 'Image' : 'Files') || 'New chat';
          const newChatRecord: IStoredChat = {
            id: newId,
            projectId: currentProjectId,
            title,
            messages: [],
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString()
          };
          chatIdToUse = newId;
          setCurrentChatIdState(newId);
          persistChatId(newId);
          if (useSharePoint && spConfig) {
            saveChatToLibrary(spConfig, newChatRecord).then(() => {
              setAllChats((prev) => [...prev, newChatRecord]);
            }).catch(() => {});
          } else {
            const all = getAllChats();
            all.push(newChatRecord);
            saveAllChats(all);
            setChats(getChatsForProject(currentProjectId));
          }
        }
        setMessages((prev) => [...prev, userMessage]);
      }

      const assistantPlaceholder: IChatMessage = {
        id: generateId(),
        role: 'assistant',
        content: '',
        timestamp: new Date()
      };
      setMessages((prev) => [...prev, assistantPlaceholder]);
      setLoading(true);
      abortRef.current = new AbortController();

      try {
        const apiMessages = buildOpenAIMessages(userText, imageDataUrls, fileTexts, otherFileNames, options?.regenerate ? { forRegenerate: true } : undefined);
        const res = await fetch(OPENAI_API_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${key}` },
          body: JSON.stringify({
            model,
            messages: apiMessages,
            max_tokens: 1024,
            stream: true
          }),
          signal: abortRef.current.signal
        });

        if (!res.ok) {
          const errBody = await res.text();
          let errMsg = `Request failed (${res.status})`;
          try {
            const j = JSON.parse(errBody);
            if (j.error?.message) errMsg = j.error.message;
          } catch {
            if (errBody) errMsg = errBody.slice(0, 200);
          }
          throw new Error(errMsg);
        }

        const reader = res.body?.getReader();
        if (!reader) throw new Error('No response body');
        const decoder = new TextDecoder();
        let buffer = '';
        let fullContent = '';

        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          buffer += decoder.decode(value, { stream: true });
          const lines = buffer.split('\n');
          buffer = lines.pop() || '';
          for (const line of lines) {
            if (line.startsWith('data: ')) {
              const data = line.slice(6);
              if (data === '[DONE]') continue;
              try {
                const parsed = JSON.parse(data);
                const delta = parsed.choices?.[0]?.delta?.content;
                if (delta) {
                  fullContent += delta;
                  setMessages((prev) =>
                    prev.map((m) => (m.id === assistantPlaceholder.id ? { ...m, content: fullContent } : m))
                  );
                }
              } catch {
                // skip malformed chunk
              }
            }
          }
        }

        if (!fullContent.trim()) fullContent = 'No response.';
        setMessages((prev) =>
          prev.map((m) => (m.id === assistantPlaceholder.id ? { ...m, content: fullContent } : m))
        );
      } catch (err) {
        if ((err as Error).name === 'AbortError') return;
        const errMsg = err instanceof Error ? err.message : 'Something went wrong.';
        setMessages((prev) =>
          prev.map((m) =>
            m.id === assistantPlaceholder.id ? { ...m, content: `Error: ${errMsg}`, role: 'assistant' } : m
          )
        );
        setError(errMsg);
      } finally {
        setLoading(false);
        abortRef.current = null;
      }
    },
    [input, attachments, loading, apiKey, messages, model, buildOpenAIMessages, currentChatId, currentProjectId]
  );

  const handleKeyDown = (e: React.KeyboardEvent): void => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  };

  const copyToClipboard = (text: string): void => {
    navigator.clipboard.writeText(text).catch(() => {});
  };

  const lastUserMessage = messages.length >= 2 && messages[messages.length - 2]?.role === 'user';
  const showApiKeyWarning = !apiKey?.trim();
  const canSend = (input.trim().length > 0 || attachments.length > 0) && !loading;

  const suggestedModel = getSuggestedModelValue(input, attachments.length > 0);
  const suggestedOption = MODEL_OPTIONS.find((o) => o.value === suggestedModel);
  const showModelSuggestion = canSend && suggestedOption && suggestedModel !== model && !loading;

  const handleAddProject = (): void => {
    const name = window.prompt('Project name', 'New project');
    if (!name?.trim()) return;
    const newProject: IProject = { id: generateId(), name: name.trim(), createdAt: new Date().toISOString() };
    const next = [...projects, newProject];
    if (useSharePoint && spConfig) {
      saveProjectsToLibrary(spConfig, next).then(() => {
        setProjects(next);
        setCurrentProjectId(newProject.id);
      }).catch((e) => setError(e instanceof Error ? e.message : 'Failed to save project'));
    } else {
      saveProjects(next);
      setProjects(next);
      setCurrentProjectId(newProject.id);
    }
  };

  const handleSelectChat = (chat: IStoredChat): void => {
    setSidebarMenuOpen(null);
    setDropdownAnchor(null);
    setCurrentChatId(chat.id);
    if (isMobile) setSidebarOpen(false);
  };

  React.useEffect(() => {
    if (!sidebarMenuOpen) return;
    const onDocClick = (e: MouseEvent): void => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) {
        setSidebarMenuOpen(null);
        setDropdownAnchor(null);
      }
    };
    document.addEventListener('click', onDocClick);
    return () => document.removeEventListener('click', onDocClick);
  }, [sidebarMenuOpen]);

  const handleRenameProject = (p: IProject): void => {
    setSidebarMenuOpen(null);
    setDropdownAnchor(null);
    const name = window.prompt('Rename project', p.name);
    if (!name?.trim() || name === p.name) return;
    const next = projects.map((x) => (x.id === p.id ? { ...x, name: name.trim() } : x));
    if (useSharePoint && spConfig) {
      saveProjectsToLibrary(spConfig, next).then(() => setProjects(next)).catch((e) => setError(e instanceof Error ? e.message : 'Failed to rename'));
    } else {
      saveProjects(next);
      setProjects(next);
    }
  };

  const handleDeleteProject = (p: IProject): void => {
    setSidebarMenuOpen(null);
    setDropdownAnchor(null);
    if (!window.confirm(`Delete project "${p.name}" and all its chats?`)) return;
    const projectChats = chats.filter((c) => c.projectId === p.id);
    if (useSharePoint && spConfig) {
      Promise.all(projectChats.map((c) => deleteChatFromLibrary(spConfig, c.id)))
        .then(() => saveProjectsToLibrary(spConfig, projects.filter((x) => x.id !== p.id)))
        .then(() => {
          setProjects((prev) => prev.filter((x) => x.id !== p.id));
          setAllChats((prev) => prev.filter((c) => c.projectId !== p.id));
          setChats((prev) => prev.filter((c) => c.projectId !== p.id));
          if (currentProjectId === p.id) {
            const remaining = projects.filter((x) => x.id !== p.id);
            const nextAll = allChats.filter((c) => c.projectId !== p.id);
            if (remaining.length > 0) {
              setCurrentProjectId(remaining[0].id);
              const firstInNew = nextAll.find((c) => c.projectId === remaining[0].id);
              setCurrentChatId(firstInNew?.id ?? null);
            } else {
              setCurrentProjectId(null);
              setCurrentChatId(null);
              setMessages([]);
            }
          }
        })
        .catch((e) => setError(e instanceof Error ? e.message : 'Failed to delete project'));
    } else {
      const nextProjects = projects.filter((x) => x.id !== p.id);
      const all = getAllChats().filter((c) => c.projectId !== p.id);
      saveProjects(nextProjects);
      saveAllChats(all);
      setProjects(nextProjects);
      setAllChats(all);
      setChats((prev) => prev.filter((c) => c.projectId !== p.id));
      if (currentProjectId === p.id) {
        if (nextProjects.length > 0) {
          setCurrentProjectId(nextProjects[0].id);
          const firstInNew = all.find((c) => c.projectId === nextProjects[0].id);
          setCurrentChatId(firstInNew?.id ?? null);
        } else {
          setCurrentProjectId(null);
          setCurrentChatId(null);
          setMessages([]);
        }
      }
    }
  };

  const handleRenameChat = (c: IStoredChat): void => {
    setSidebarMenuOpen(null);
    setDropdownAnchor(null);
    const title = window.prompt('Rename chat', c.title || 'New chat');
    if (title === null || title === (c.title || 'New chat')) return;
    const updated = { ...c, title: title.trim() || 'New chat', updatedAt: new Date().toISOString() };
    if (useSharePoint && spConfig) {
      saveChatToLibrary(spConfig, updated).then(() => {
        setAllChats((prev) => prev.map((x) => (x.id === c.id ? updated : x)));
      }).catch((e) => setError(e instanceof Error ? e.message : 'Failed to rename chat'));
    } else {
      const all = getAllChats();
      const idx = all.findIndex((x) => x.id === c.id);
      if (idx >= 0) {
        all[idx] = updated;
        saveAllChats(all);
        setAllChats([...all]);
        setChats(getChatsForProject(currentProjectId));
      }
    }
  };

  const handleDeleteChat = (chat: IStoredChat): void => {
    setSidebarMenuOpen(null);
    setDropdownAnchor(null);
    if (!window.confirm('Delete this chat?')) return;
    if (useSharePoint && spConfig) {
      deleteChatFromLibrary(spConfig, chat.id)
        .then(() => {
          setAllChats((prev) => prev.filter((c) => c.id !== chat.id));
          setChats((prev) => prev.filter((c) => c.id !== chat.id));
          if (currentChatId === chat.id) {
            setCurrentChatId(null);
            setMessages([]);
            const rest = chats.filter((c) => c.id !== chat.id);
            if (rest.length > 0) handleSelectChat(rest[0]);
          }
        })
        .catch((e) => setError(e instanceof Error ? e.message : 'Failed to delete chat'));
    } else {
      const all = getAllChats().filter((x) => x.id !== chat.id);
      saveAllChats(all);
      setAllChats(all);
      setChats(getChatsForProject(currentProjectId));
      if (currentChatId === chat.id) {
        setCurrentChatId(null);
        setMessages([]);
        const rest = all.filter((c) => c.projectId === currentProjectId);
        if (rest.length > 0) setCurrentChatId(rest[0].id);
      } else {
        setChats((prev) => prev.filter((c) => c.id !== chat.id));
      }
    }
  };

  return (
    <section className={styles.chat}>
      {sidebarOpen && isMobile && (
        <div
          className={styles.sidebarBackdrop}
          onClick={() => setSidebarOpen(false)}
          role="button"
          tabIndex={0}
          aria-label="Close menu"
          onKeyDown={(e) => e.key === 'Enter' && setSidebarOpen(false)}
        />
      )}
      {sidebarOpen && (
        <aside className={styles.sidebar}>
          {storageLoading && (
            <div className={styles.sidebarLoading}>Loading from library…</div>
          )}
          {useSharePoint && !storageLoading && (
            <div className={styles.sidebarStorageHint}>Stored in: {docsLibraryName}</div>
          )}
          <div className={styles.sidebarSection}>
            <div className={styles.sidebarSectionHeader}>
              <span>Projects</span>
              <button type="button" className={styles.sidebarBtn} onClick={handleAddProject} title="New project" disabled={storageLoading}>+</button>
            </div>
            <ul className={styles.sidebarList}>
              {projects.map((p) => (
                <li key={p.id} className={styles.sidebarItemRow}>
                  <button
                    type="button"
                    className={styles.sidebarItem + (p.id === currentProjectId ? ' ' + styles.sidebarItemActive : '')}
                    onClick={() => { setSidebarMenuOpen(null); setDropdownAnchor(null); setCurrentProjectId(p.id); if (isMobile) setSidebarOpen(false); }}
                  >
                    {p.name}
                  </button>
                  <button
                    type="button"
                    className={styles.sidebarItemMenu}
                    onClick={(e) => {
                      e.stopPropagation();
                      const rect = e.currentTarget.getBoundingClientRect();
                      if (sidebarMenuOpen === 'project-' + p.id) {
                        setSidebarMenuOpen(null);
                        setDropdownAnchor(null);
                      } else {
                        setDropdownAnchor({ top: rect.top, left: rect.left, bottom: rect.bottom });
                        setSidebarMenuOpen('project-' + p.id);
                      }
                    }}
                    title="Project options"
                    aria-label="Project options"
                  >
                    &#8942;
                  </button>
                </li>
              ))}
            </ul>
          </div>
          <div className={styles.sidebarSection}>
            <div className={styles.sidebarSectionHeader}>
              <span>Chats</span>
              <button type="button" className={styles.sidebarBtn} onClick={newChat} title="New chat" disabled={storageLoading}>+</button>
            </div>
            <ul className={styles.sidebarList}>
              {chats.map((c) => (
                <li key={c.id} className={styles.sidebarItemRow}>
                  <button
                    type="button"
                    className={styles.sidebarItem + (c.id === currentChatId ? ' ' + styles.sidebarItemActive : '')}
                    onClick={() => handleSelectChat(c)}
                  >
                    {c.title || 'New chat'}
                  </button>
                  <button
                    type="button"
                    className={styles.sidebarItemMenu}
                    onClick={(e) => {
                      e.stopPropagation();
                      const rect = e.currentTarget.getBoundingClientRect();
                      if (sidebarMenuOpen === 'chat-' + c.id) {
                        setSidebarMenuOpen(null);
                        setDropdownAnchor(null);
                      } else {
                        setDropdownAnchor({ top: rect.top, left: rect.left, bottom: rect.bottom });
                        setSidebarMenuOpen('chat-' + c.id);
                      }
                    }}
                    title="Chat options"
                    aria-label="Chat options"
                  >
                    &#8942;
                  </button>
                </li>
              ))}
            </ul>
          </div>
        </aside>
      )}
      {sidebarMenuOpen && dropdownAnchor && (() => {
        const isProject = sidebarMenuOpen.startsWith('project-');
        const id = isProject ? sidebarMenuOpen.slice(8) : sidebarMenuOpen.slice(5);
        const project = isProject ? projects.find((p) => p.id === id) : null;
        const chat = !isProject ? allChats.find((c) => c.id === id) : null;
        return (
          <div
            ref={dropdownRef}
            className={styles.sidebarDropdownFixed}
            style={{ left: dropdownAnchor.left, top: dropdownAnchor.bottom + 4 }}
          >
            {project && (
              <>
                <button type="button" onClick={() => handleRenameProject(project)}>Rename</button>
                <button type="button" onClick={() => handleDeleteProject(project)}>Delete</button>
              </>
            )}
            {chat && (
              <>
                <button type="button" onClick={() => handleRenameChat(chat)}>Rename</button>
                <button type="button" onClick={() => handleDeleteChat(chat)}>Delete</button>
              </>
            )}
          </div>
        );
      })()}
      <div className={styles.main}>
      <div className={styles.header}>
        <button type="button" className={styles.sidebarToggle} onClick={() => setSidebarOpen((o) => !o)} title={sidebarOpen ? 'Hide sidebar' : 'Show sidebar'} aria-label={sidebarOpen ? 'Hide sidebar' : 'Show sidebar'}>
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M3 6h18M3 12h18M3 18h18"/></svg>
        </button>
        <span>Chat</span>
        <div className={styles.headerActions}>
          <label className={styles.modelLabel}>
            <span className={styles.modelLabelText}>Model</span>
            <select
              className={styles.modelSelect}
              value={model}
              onChange={(e) => setModel(e.target.value)}
              disabled={loading}
              title="OpenAI model"
              aria-label="Select model"
            >
              {MODEL_OPTIONS.map((opt) => (
                <option key={opt.value} value={opt.value}>{opt.label}</option>
              ))}
            </select>
          </label>
          <button type="button" className={styles.headerBtn} onClick={newChat} title="New chat">
            New chat
          </button>
        </div>
      </div>

      {showApiKeyWarning && (
        <div className={styles.apiKeyWarning}>
          OpenAI API key is not set. Add it in web part properties (edit page → select web part → pencil icon).
        </div>
      )}

      <div className={styles.messages}>
        {messages.length === 0 && (
          <div className={styles.welcome}>
            Type a message or attach files below. Supported: PDF, Word (.docx), Excel (.xlsx, .xls), text (.txt, .md, .csv, .json, .log, .yml, .rtf, .sql, .ini, .env, etc.), and images. You can attach multiple files.
          </div>
        )}
        {messages.map((msg) => (
          <div
            key={msg.id}
            className={`${styles.message} ${styles[msg.role]} ${msg.content.indexOf('Error:') === 0 ? styles.error : ''}`}
          >
            <div className={styles.messageBody}>
              {msg.role === 'user' ? (
                <>
                  {msg.content}
                  {msg.imageDataUrls?.length ? (
                    <div className={styles.messageImages}>
                      {msg.imageDataUrls.map((url, i) => (
                        <img key={i} src={url} alt="" className={styles.messageImg} />
                      ))}
                    </div>
                  ) : null}
                  {(msg.fileTexts?.length || msg.otherFileNames?.length) ? (
                    <div className={styles.messageFiles}>
                      {msg.fileTexts?.map((f, i) => (
                        <div key={`t-${i}`} className={styles.messageFileChip}>{f.name}</div>
                      ))}
                      {msg.otherFileNames?.map((name, i) => (
                        <div key={`o-${i}`} className={styles.messageFileChipUnsupported}>{name}</div>
                      ))}
                    </div>
                  ) : null}
                </>
              ) : (
                <div className={styles.assistantContent}>
                  {msg.content ? renderMarkdown(msg.content) : loading ? 'Thinking…' : ''}
                  {msg.content && msg.role === 'assistant' && (
                    <button
                      type="button"
                      className={styles.copyBtn}
                      onClick={() => copyToClipboard(msg.content)}
                      title="Copy"
                      aria-label="Copy"
                    >
                      <svg className={styles.copyBtnIcon} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                        <rect x="9" y="9" width="13" height="13" rx="2" ry="2" />
                        <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
                      </svg>
                    </button>
                  )}
                </div>
              )}
            </div>
          </div>
        ))}
        <div ref={messagesEndRef} />
      </div>

      {attachments.length > 0 && (
        <div className={styles.attachmentBar}>
          {attachments.map((a, i) => (
            <div key={i} className={styles.attachmentPreview}>
              {a.dataUrl ? (
                <>
                  <img src={a.dataUrl} alt="" />
                  <span className={styles.attachmentFileName} title={a.file.name}>{a.file.name}</span>
                </>
              ) : (
                <>
                  <span className={styles.attachmentIconWrap}>
                    <FileTypeIcon type={getFileThumbType(a.file)} className={styles.attachmentFileIcon} />
                  </span>
                  <span className={styles.attachmentDocName} title={a.file.name}>
                    {a.file.name}
                    {a.textContent !== undefined ? (
                      <span className={styles.attachmentBadge}> text ready</span>
                    ) : a.extractionError ? (
                      <span className={styles.attachmentBadgeError}> extraction failed</span>
                    ) : null}
                  </span>
                </>
              )}
              <button type="button" className={styles.attachmentRemove} onClick={() => removeAttachment(i)} aria-label="Remove">×</button>
            </div>
          ))}
        </div>
      )}

      {showModelSuggestion && (
        <div className={styles.modelSuggestion}>
          <span className={styles.modelSuggestionText}>
            {suggestedModel === MODEL_FOR_COMPLEX
              ? 'This looks like a complex question. Consider a stronger model for better analysis.'
              : 'This looks straightforward. You can use a faster, cheaper model.'}
          </span>
          <button
            type="button"
            className={styles.modelSuggestionBtn}
            onClick={() => setModel(suggestedModel)}
            title={`Use ${suggestedOption?.label ?? suggestedModel}`}
          >
            Use {suggestedOption?.label ?? suggestedModel}
          </button>
        </div>
      )}

      <div className={styles.inputArea}>
        <div className={styles.attachWrap + (loading ? ' ' + styles.attachWrapDisabled : '')}>
          <input
            ref={fileInputRef}
            type="file"
            multiple
            className={styles.fileInput}
            onChange={onFileChange}
            aria-label="Upload files"
            disabled={loading}
            title="Upload files (images or documents)"
          />
          <span className={styles.attachButton} aria-hidden>
            <svg className={styles.attachIcon} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
              <path d="M21.44 11.05l-9.19 9.19a6 6 0 01-8.49-8.49l9.19-9.19a4 4 0 015.66 5.66l-9.2 9.19a2 2 0 01-2.83-2.83l8.49-8.48" />
            </svg>
          </span>
        </div>
        <textarea
          ref={textareaRef}
          className={styles.input}
          placeholder="Type a message or attach files..."
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={handleKeyDown}
          rows={1}
          disabled={loading}
        />
        <button
          type="button"
          className={styles.sendButton}
          onClick={() => sendMessage()}
          disabled={!canSend}
        >
          {loading ? <span className={styles.loading} aria-hidden /> : 'Send'}
        </button>
      </div>

      {lastUserMessage && messages.length >= 2 && !loading && (
        <div className={styles.regenerateBar}>
          <button type="button" className={styles.regenerateBtn} onClick={() => sendMessage({ regenerate: true })}>
            Regenerate response
          </button>
        </div>
      )}
      </div>
    </section>
  );
};

export default Chat;
