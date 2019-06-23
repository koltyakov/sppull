import { IFileBasicMetadata } from './';

export interface IContent {
  files: IFile[];
  folders: IFolder[];
}

export interface IFile extends IFileBasicMetadata {
  [prop: string]: any;
}

export interface IFolder {
  [prop: string]: any;
}
