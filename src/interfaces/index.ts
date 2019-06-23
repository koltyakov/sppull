import { IAuthOptions } from 'node-sp-auth';
import RestAPI from '../api';

export interface IFileBasicMetadata {
  ServerRelativeUrl: string;
  Name?: string;
  UniqueID?: string;
  ID?: string;
  Length?: number;
  TimeCreated?: string;
  TimeLastModified?: string;
  SavedToLocalPath?: string;
}

export interface ISPPullContext {
  siteUrl: string;
  creds: IAuthOptions;
}

export interface ISPPullOptions {
  spHostName?: string;
  spRelativeBase?: string;
  spRootFolder?: string;
  spBaseFolder?: string;
  fileRegExp?: RegExp;
  dlRootFolder?: string;
  recursive?: boolean;
  foderStructureOnly?: boolean;
  createEmptyFolders?: boolean;
  metaFields?: string[];
  restCondition?: string;
  muteConsole?: boolean;
  camlCondition?: string;
  spDocLibUrl?: string;
  omitFolderPath?: string;
  strictObjects?: string[];
}

export interface ICtx {
  context: ISPPullContext;
  options: ISPPullOptions;
  api: RestAPI;
}
