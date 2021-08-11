import { IFileBasicMetadata } from './';

export interface IContent {
  files: IFileMetadata[];
  folders: IFolderMetadata[];
}

export interface IFileMetadata extends IFileBasicMetadata {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  [prop: string]: any;
}

export interface IFolderMetadata {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  [prop: string]: any;
}
