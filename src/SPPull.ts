import * as path from 'path';
import * as mkdirp from 'mkdirp';
import * as colors from 'colors';
import * as readline from 'readline';
import { URL } from 'url';

import RestAPI from './api';
import { ISPPullOptions, ISPPullContext, IFileBasicMetadata, ICtx } from './interfaces';
import { IContent, IFolder } from './interfaces/content';

export class Download {

  public sppull = (context: ISPPullContext, options: ISPPullOptions) => {

    // Apply defeults
    options = this.initOptions(context, options);

    const ctx: ICtx = {
      context,
      options,
      api: new RestAPI(context, options)
    };

    if (
      typeof options.camlCondition !== 'undefined' && options.camlCondition !== '' &&
      typeof options.spDocLibUrl !== 'undefined' && options.spDocLibUrl !== ''
    ) {
      return this.runDownloadCamlObjects(ctx);
    } else {
      if (typeof options.strictObjects !== 'undefined' && Array.isArray([options.strictObjects])) {
        options.strictObjects.forEach((strictObject, i) => {
          if (typeof strictObject === 'string') {
            if (strictObject.indexOf(options.spRootFolder) !== 0) {
              strictObject = (options.spRootFolder + '/' + strictObject).replace(/\/\//g, '/');
            }
            options.strictObjects[i] = strictObject;
          }
        });
        return this.runDownloadStrictObjects(ctx);
      } else {
        if (!options.foderStructureOnly) {
          if (options.recursive) {
            return this.runDownloadFilesRecursively(ctx);
          } else {
            return this.runDownloadFilesFlat(ctx);
          }
        } else {
          return this.runCreateFoldersRecursively(ctx);
        }
      }

    }
  }

  private createFolder(ctx: ICtx, spFolderPath: string, downloadRoot: string): Promise<string> {
    return new Promise((resolve, reject) => {

      const spBaseFolderRegEx = new RegExp(decodeURIComponent(ctx.options.spBaseFolder), 'gi');
      let spFolderPathRelative = decodeURIComponent(spFolderPath);
      if (['', '/'].indexOf(ctx.options.spBaseFolder) === -1) {
        spFolderPathRelative = decodeURIComponent(spFolderPath).replace(spBaseFolderRegEx, '');
      }

      const saveFolderPath: string = path.join(downloadRoot, spFolderPathRelative);

      mkdirp(saveFolderPath)
        .then(() => resolve(saveFolderPath))
        .catch((err) => {
          console.log(colors.red.bold('Error creating folder ' + '`' + saveFolderPath + ' `'), colors.red(err));
          reject(err);
        });
    });
  }

  // Queues >>>>
  private async createFoldersQueue(ctx: ICtx, foldersList: IFolder[], index: number = 0): Promise<IFolder[]> {
    const spFolderPath: string = foldersList[index].ServerRelativeUrl;
    const downloadRoot = ctx.options.dlRootFolder;

    if (!ctx.options.muteConsole) {
      readline.clearLine(process.stdout, 0);
      readline.cursorTo(process.stdout, 0, undefined);
      process.stdout.write(colors.green.bold('Creating folders: ') + (index + 1) + ' out of ' + foldersList.length);
    }

    const localFolderPath = await this.createFolder(ctx, spFolderPath, downloadRoot);
    foldersList[index].SavedToLocalPath = localFolderPath;
    index += 1;
    if (index < foldersList.length) {
      return this.createFoldersQueue(ctx, foldersList, index);
    } else {
      if (!ctx.options.muteConsole) {
        process.stdout.write('\n');
      }
      return foldersList;
    }
  }

  private async downloadFilesQueue(ctx: ICtx, filesList: IFileBasicMetadata[], index: number = 0): Promise<IFileBasicMetadata[]> {
    const spFilePath = filesList[index].ServerRelativeUrl;
    if (!ctx.options.muteConsole) {
      readline.clearLine(process.stdout, 0);
      readline.cursorTo(process.stdout, 0, undefined);
      process.stdout.write(colors.green.bold('Downloading files: ') + (index + 1) + ' out of ' + filesList.length);
    }
    await ctx.api.downloadFile(spFilePath, filesList[index])
      .then((localFilePath) => {
        filesList[index].SavedToLocalPath = localFilePath;
      })
      .catch((ex) => {
        const err = ex.message || ex;
        console.log('\n', colors.red.bold('\nError in operations.downloadFile:'), colors.red(err));
        filesList[index].Error = err;
      });
    index += 1;
    if (index < filesList.length) {
      return this.downloadFilesQueue(ctx, filesList, index);
    } else {
      if (!ctx.options.muteConsole) {
        process.stdout.write('\n');
      }
      return filesList;
    }
  }

  private async getStructureRecursive(ctx: ICtx, root: boolean = true, foldersQueue: any[] = [], filesList: any[] = []): Promise<IContent> {
    let exitQueue = true;
    if (typeof ctx.options.spRootFolder === 'undefined' || ctx.options.spRootFolder === '') {
      throw Error('The `spRootFolder` property should be provided in options.');
    }
    let spRootFolder;

    if (foldersQueue.length === 0) {
      spRootFolder = ctx.options.spRootFolder;
      exitQueue = !root; // false;
    } else {
      foldersQueue.some((fi) => {
        if (typeof fi.processed === 'undefined') {
          fi.processed = false;
        }
        if (!fi.processed) {
          spRootFolder = fi.serverRelativeUrl;
          fi.processed = true;
          exitQueue = false;
          return true;
        }
        return false;
      });
    }

    if (!exitQueue) {
      let cntInQueue = 0;
      foldersQueue.forEach((folder) => {
        if (folder.processed) {
          cntInQueue += 1;
        }
      });

      if (!ctx.options.muteConsole) {
        readline.clearLine(process.stdout, 0);
        readline.cursorTo(process.stdout, 0, undefined);
        process.stdout.write(colors.green.bold('Folders proceeding: ') + cntInQueue + ' out of ' + foldersQueue.length + colors.gray(' [recursive scanning...]'));
      }

      const results = await ctx.api.getFolderContent(spRootFolder);
      (results.folders || []).forEach((folder) => {
        const folderElement = {
          folder,
          serverRelativeUrl: folder.ServerRelativeUrl,
          processed: false
        };
        foldersQueue.push(folderElement);
      });
      filesList = filesList.concat(results.files || []);
      return this.getStructureRecursive(ctx, false, foldersQueue, filesList);

    } else {
      if (!ctx.options.muteConsole) {
        process.stdout.write('\n');
      }
      const foldersList = foldersQueue.map((folder) => {
        return folder.folder;
      });
      return {
        files: filesList,
        folders: foldersList
      };
    }
  }
  // <<<< Queues

  // Runners >>>>
  private async runCreateFoldersRecursively(ctx: ICtx) {
    const data = await this.getStructureRecursive(ctx);
    if ((data.folders || []).length > 0) {
      return this.createFoldersQueue(ctx, data.folders, 0);
    }
    return [];
  }

  private async runDownloadFilesRecursively(ctx: ICtx) {
    const data = await this.getStructureRecursive(ctx);
    if (ctx.options.createEmptyFolders) {
      const folders = data.folders || [];
      if (folders.length > 0) {
        await this.createFoldersQueue(ctx, folders, 0);
      }
    }
    return this.downloadMyFilesHandler(ctx, data);
  }

  private async runDownloadFilesFlat(ctx: ICtx) {
    if (typeof ctx.options.spRootFolder === 'undefined' || ctx.options.spRootFolder === '') {
      throw Error('The `spRootFolder` property should be provided in options.');
    }
    const data = await ctx.api.getFolderContent(ctx.options.spRootFolder);
    if (ctx.options.createEmptyFolders) {
      if ((data.folders || []).length > 0) {
        await this.createFoldersQueue(ctx, data.folders, 0);
      }
    }
    return this.downloadMyFilesHandler(ctx, data);
  }

  private async runDownloadStrictObjects(ctx: ICtx) {
    const filesList: IFileBasicMetadata[] = ctx.options.strictObjects
      .filter((d) => d.split('/').slice(-1)[0].indexOf('.') !== -1)
      .map((ServerRelativeUrl) => ({ ServerRelativeUrl }));

    if (filesList.length > 0) {
      return this.downloadFilesQueue(ctx, filesList, 0);
    }
    return [];
  }

  private async runDownloadCamlObjects(ctx: ICtx) {
    const data = await ctx.api.getContentWithCaml();
    if (ctx.options.createEmptyFolders) {
      if ((data.folders || []).length > 0) {
        await this.createFoldersQueue(ctx, data.folders, 0);
      }
    }
    return this.downloadMyFilesHandler(ctx, data);
  }
  // <<<< Runners

  private async downloadMyFilesHandler(ctx: ICtx, data: IContent) {
    let files: { ServerRelativeUrl: string }[] = data.files || [];
    const { fileRegExp } = ctx.options;
    if (typeof fileRegExp === 'object' && typeof fileRegExp.test === 'function') {
      files = files.filter(f => fileRegExp.test(f.ServerRelativeUrl));
    }
    if (files.length > 0) {
      return this.downloadFilesQueue(ctx, files, 0);
    }
    return [];
  }

  private initOptions(context: ISPPullContext, options: ISPPullOptions): ISPPullOptions {
    if (typeof context.creds === 'undefined') {
      context.creds = { ...(context as any) };
    }

    const url = new URL(context.siteUrl);

    options.spHostName = url.hostname;
    options.spRelativeBase = url.pathname;

    if (options.spRootFolder) {
      if (options.spRootFolder.indexOf(options.spRelativeBase) !== 0) {
        options.spRootFolder = (options.spRelativeBase + '/' + options.spRootFolder).replace(/\/\//g, '/');
      } else {
        if (options.spRootFolder.charAt(0) !== '/') {
          options.spRootFolder = '/' + options.spRootFolder;
        }
      }
    }

    if (options.spBaseFolder) {
      if (options.spBaseFolder.indexOf(options.spRelativeBase) !== 0) {
        options.spBaseFolder = (options.spRelativeBase + '/' + options.spBaseFolder).replace(/\/\//g, '/');
      }
    } else {
      options.spBaseFolder = options.spRootFolder;
    }

    options.dlRootFolder = options.dlRootFolder || './Downloads';
    options.recursive = typeof options.recursive !== 'undefined' ? options.recursive : true;
    options.foderStructureOnly = typeof options.foderStructureOnly !== 'undefined' ? options.foderStructureOnly : false;
    options.createEmptyFolders = typeof options.createEmptyFolders !== 'undefined' ? options.createEmptyFolders : true;
    options.metaFields = options.metaFields || [];

    if (options.spDocLibUrl) {
      if (options.spDocLibUrl.indexOf(options.spRelativeBase) !== 0) {
        options.spDocLibUrl = (options.spRelativeBase + '/' + options.spDocLibUrl).replace(/\/\//g, '/');
      } else {
        if (options.spDocLibUrl.charAt(0) !== '/') {
          options.spDocLibUrl = '/' + options.spDocLibUrl;
        }
      }
      options.spDocLibUrl = encodeURIComponent(options.spDocLibUrl);
    }

    if (typeof options.muteConsole === 'undefined') {
      options.muteConsole = false;
    }

    return options;
  }

}

export { ISPPullOptions, ISPPullContext } from './interfaces';
