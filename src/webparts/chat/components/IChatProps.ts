import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IChatProps {
  apiKey: string;
  context: WebPartContext;
  /** Document library title on this site for storing chats (filtered by user). Empty = use browser storage. */
  docsLibraryName: string;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  userDisplayName: string;
}

export interface IChatMessage {
  id: string;
  role: 'user' | 'assistant';
  content: string;
  timestamp: Date;
  /** Data URLs for images attached to this message (user only) */
  imageDataUrls?: string[];
  /** Extracted text from attached documents (user only) */
  fileTexts?: { name: string; text: string }[];
  /** Names of files attached but content not extracted (user only) */
  otherFileNames?: string[];
}

/** Attachment prepared for the next send: image (dataUrl) or document (textContent) */
export interface IChatAttachment {
  file: File;
  /** Present for image files */
  dataUrl?: string;
  /** Present for text/document files (extracted content) */
  textContent?: string;
  /** Set when extraction was attempted but failed (e.g. PDF in browser) */
  extractionError?: string;
}

/** Stored in localStorage */
export interface IProject {
  id: string;
  name: string;
  createdAt: string;
}

/** Message as stored (timestamp as ISO string) */
export interface IStoredMessage {
  id: string;
  role: 'user' | 'assistant';
  content: string;
  timestamp: string;
  imageDataUrls?: string[];
  fileTexts?: { name: string; text: string }[];
  otherFileNames?: string[];
}

/** Stored in localStorage */
export interface IStoredChat {
  id: string;
  projectId: string;
  title: string;
  messages: IStoredMessage[];
  createdAt: string;
  updatedAt: string;
}
