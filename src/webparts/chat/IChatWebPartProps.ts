export interface IChatWebPartProps {
  apiKey: string;
  /** Document library on this site where chats are stored (one file per chat, filtered by current user). Leave blank to use browser storage. */
  docsLibraryName: string;
}
