/**
 * Store projects and chats in a site document library. One file per chat, one file per user for projects.
 * Filter items by current user (Author).
 */

import { SPHttpClient } from '@microsoft/sp-http';
import { IProject, IStoredChat } from './IChatProps';

const PROJECTS_FILE_PREFIX = 'chat2etc-projects-';
const CHAT_FILE_PREFIX = 'chat2etc-chat-';
const FILE_EXT = '.json';

export interface IChatStorageSharePointConfig {
  spHttpClient: SPHttpClient;
  webAbsoluteUrl: string;
  libraryTitle: string;
}

let cachedUserId: number | null = null;

export async function getCurrentUserId(config: IChatStorageSharePointConfig): Promise<number> {
  if (cachedUserId !== null) return cachedUserId;
  const url = `${config.webAbsoluteUrl}/_api/web/currentuser?$select=Id`;
  const res = await config.spHttpClient.get(url, SPHttpClient.configurations.v1);
  if (!res.ok) throw new Error('Failed to get current user');
  const json = await res.json();
  cachedUserId = json.Id as number;
  return cachedUserId;
}

async function getFolderServerRelativeUrl(config: IChatStorageSharePointConfig): Promise<string> {
  const url = `${config.webAbsoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(config.libraryTitle)}')/rootFolder?$select=ServerRelativeUrl`;
  const res = await config.spHttpClient.get(url, SPHttpClient.configurations.v1);
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Library not found or no access: ${config.libraryTitle}. ${text.slice(0, 200)}`);
  }
  const json = await res.json();
  return json.ServerRelativeUrl as string;
}

/** Get list items (files) in the library created by the current user */
async function getMyFileItems(config: IChatStorageSharePointConfig): Promise<{ Name: string; ServerRelativeUrl: string }[]> {
  const userId = await getCurrentUserId(config);
  const listTitle = encodeURIComponent(config.libraryTitle);
  const url = `${config.webAbsoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$filter=AuthorId eq ${userId}&$select=Id,File/Name,File/ServerRelativeUrl&$expand=File`;
  const res = await config.spHttpClient.get(url, SPHttpClient.configurations.v1);
  if (!res.ok) return [];
  const json = await res.json();
  const items = json.value as { File?: { Name: string; ServerRelativeUrl: string } }[];
  return items
    .filter((i) => i.File?.Name)
    .map((i) => ({ Name: i.File!.Name, ServerRelativeUrl: i.File!.ServerRelativeUrl }));
}

async function getFileContent(config: IChatStorageSharePointConfig, serverRelativeUrl: string): Promise<string> {
  const encoded = encodeURIComponent(serverRelativeUrl);
  const url = `${config.webAbsoluteUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelativeUrl.replace(/'/g, "''")}')/$value`;
  const res = await config.spHttpClient.get(url, SPHttpClient.configurations.v1);
  if (!res.ok) return '';
  return await res.text();
}

export async function getProjectsFromLibrary(config: IChatStorageSharePointConfig): Promise<IProject[]> {
  const userId = await getCurrentUserId(config);
  const fileName = PROJECTS_FILE_PREFIX + userId + FILE_EXT;
  const files = await getMyFileItems(config);
  const file = files.find((f) => f.Name === fileName);
  if (!file) return [];
  const raw = await getFileContent(config, file.ServerRelativeUrl);
  if (!raw) return [];
  try {
    const list = JSON.parse(raw) as IProject[];
    return Array.isArray(list) ? list : [];
  } catch {
    return [];
  }
}

export async function saveProjectsToLibrary(config: IChatStorageSharePointConfig, projects: IProject[]): Promise<void> {
  const userId = await getCurrentUserId(config);
  const fileName = PROJECTS_FILE_PREFIX + userId + FILE_EXT;
  const folderUrl = await getFolderServerRelativeUrl(config);
  const content = JSON.stringify(projects);
  const addUrl = `${config.webAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl.replace(/'/g, "''")}')/Files/add(url='${fileName}',overwrite=true)`;
  const res = await config.spHttpClient.post(addUrl, SPHttpClient.configurations.v1, {
    body: content,
    headers: { 'Content-Type': 'application/json' }
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Failed to save projects: ${text.slice(0, 200)}`);
  }
}

export async function getChatsFromLibrary(config: IChatStorageSharePointConfig): Promise<IStoredChat[]> {
  const files = await getMyFileItems(config);
  const chatFiles = files.filter((f) => f.Name.startsWith(CHAT_FILE_PREFIX) && f.Name.endsWith(FILE_EXT));
  const chats: IStoredChat[] = [];
  for (const file of chatFiles) {
    const raw = await getFileContent(config, file.ServerRelativeUrl);
    if (!raw) continue;
    try {
      const chat = JSON.parse(raw) as IStoredChat;
      if (chat.id && chat.projectId && Array.isArray(chat.messages)) chats.push(chat);
    } catch {
      // skip invalid file
    }
  }
  return chats.sort((a, b) => new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime());
}

export async function saveChatToLibrary(config: IChatStorageSharePointConfig, chat: IStoredChat): Promise<void> {
  const folderUrl = await getFolderServerRelativeUrl(config);
  const fileName = CHAT_FILE_PREFIX + chat.id + FILE_EXT;
  const content = JSON.stringify(chat);
  const addUrl = `${config.webAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl.replace(/'/g, "''")}')/Files/add(url='${fileName}',overwrite=true)`;
  const res = await config.spHttpClient.post(addUrl, SPHttpClient.configurations.v1, {
    body: content,
    headers: { 'Content-Type': 'application/json' }
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Failed to save chat: ${text.slice(0, 200)}`);
  }
}

/** Delete a chat file from the library (for delete chat). */
export async function deleteChatFromLibrary(config: IChatStorageSharePointConfig, chatId: string): Promise<void> {
  const files = await getMyFileItems(config);
  const fileName = CHAT_FILE_PREFIX + chatId + FILE_EXT;
  const file = files.find((f) => f.Name === fileName);
  if (!file) return;
  const deleteUrl = `${config.webAbsoluteUrl}/_api/web/GetFileByServerRelativeUrl('${file.ServerRelativeUrl.replace(/'/g, "''")}')`;
  const res = await config.spHttpClient.fetch(deleteUrl, SPHttpClient.configurations.v1, { method: 'DELETE' });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Failed to delete chat: ${text.slice(0, 200)}`);
  }
}
