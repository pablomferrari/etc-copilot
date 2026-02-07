declare interface IChatWebPartStrings {
  ChatTitle: string;
  ChatDescription: string;
  SendButton: string;
  Placeholder: string;
  ApiKeyMissing: string;
  ErrorSending: string;
  Thinking: string;
}

declare module 'ChatWebPartStrings' {
  const strings: IChatWebPartStrings;
  export = strings;
}
