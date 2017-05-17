import { IAuthOptions } from 'node-sp-auth';

export interface ISPPullContext {
    siteUrl: string;
    creds: IAuthOptions;
}

export interface ISPPullOptions {
    spHostName?: string;
    spRelativeBase?: string;
    spRootFolder?: string;
    spBaseFolder?: string;
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
