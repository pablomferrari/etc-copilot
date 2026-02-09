/**
 * Projects and chats persisted in localStorage (per origin).
 */

import { IProject, IStoredChat, IStoredMessage } from './IChatProps';
import { IChatMessage } from './IChatProps';

const KEY_PROJECTS = 'chat2etc_projects';
const KEY_CHATS = 'chat2etc_chats';
const KEY_CURRENT_PROJECT = 'chat2etc_currentProject';
const KEY_CURRENT_CHAT = 'chat2etc_currentChat';
const KEY_SAVED_PROMPTS = 'chat2etc_savedPrompts';

export interface ISavedPrompt {
  id: string;
  name: string;
  text: string;
}

export function getSavedPrompts(): ISavedPrompt[] {
  try {
    const raw = localStorage.getItem(KEY_SAVED_PROMPTS);
    if (!raw) return [];
    const list = JSON.parse(raw) as ISavedPrompt[];
    return Array.isArray(list) ? list : [];
  } catch {
    return [];
  }
}

export function saveSavedPrompts(prompts: ISavedPrompt[]): void {
  try {
    localStorage.setItem(KEY_SAVED_PROMPTS, JSON.stringify(prompts));
  } catch {
    // ignore
  }
}

export function getProjects(): IProject[] {
  try {
    const raw = localStorage.getItem(KEY_PROJECTS);
    if (!raw) return [];
    const list = JSON.parse(raw) as IProject[];
    return Array.isArray(list) ? list : [];
  } catch {
    return [];
  }
}

export function saveProjects(projects: IProject[]): void {
  try {
    localStorage.setItem(KEY_PROJECTS, JSON.stringify(projects));
  } catch {
    // ignore
  }
}

export function getAllChats(): IStoredChat[] {
  try {
    const raw = localStorage.getItem(KEY_CHATS);
    if (!raw) return [];
    const list = JSON.parse(raw) as IStoredChat[];
    return Array.isArray(list) ? list : [];
  } catch {
    return [];
  }
}

export function saveAllChats(chats: IStoredChat[]): void {
  try {
    localStorage.setItem(KEY_CHATS, JSON.stringify(chats));
  } catch {
    // ignore
  }
}

export function getChatsForProject(projectId: string): IStoredChat[] {
  return getAllChats().filter((c) => c.projectId === projectId).sort((a, b) => new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime());
}

export function getCurrentProjectId(): string | null {
  return localStorage.getItem(KEY_CURRENT_PROJECT);
}

export function setCurrentProjectId(id: string | null): void {
  if (id) localStorage.setItem(KEY_CURRENT_PROJECT, id);
  else localStorage.removeItem(KEY_CURRENT_PROJECT);
}

export function getCurrentChatId(): string | null {
  return localStorage.getItem(KEY_CURRENT_CHAT);
}

export function setCurrentChatId(id: string | null): void {
  if (id) localStorage.setItem(KEY_CURRENT_CHAT, id);
  else localStorage.removeItem(KEY_CURRENT_CHAT);
}

export function messageToStored(m: IChatMessage): IStoredMessage {
  return {
    id: m.id,
    role: m.role,
    content: m.content,
    timestamp: typeof m.timestamp === 'string' ? m.timestamp : m.timestamp.toISOString(),
    imageDataUrls: m.imageDataUrls,
    fileTexts: m.fileTexts,
    otherFileNames: m.otherFileNames
  };
}

export function storedToMessage(s: IStoredMessage): IChatMessage {
  return {
    id: s.id,
    role: s.role,
    content: s.content,
    timestamp: new Date(s.timestamp),
    imageDataUrls: s.imageDataUrls,
    fileTexts: s.fileTexts,
    otherFileNames: s.otherFileNames
  };
}
