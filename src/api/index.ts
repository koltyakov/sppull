import * as fs from 'fs';
import * as path from 'path';
import * as https from 'https';
import * as mkdirp from 'mkdirp';
import * as sprequest from 'sp-request';
import * as request from 'request';
import { getAuth } from 'node-sp-auth';

import { Utils } from './../utils';
import { ISPPullOptions, ISPPullContext, IFileBasicMetadata } from '../interfaces';
import { IContent, IFile, IFolder } from '../interfaces/content';

export default class RestAPI {

  private context: ISPPullContext;
  private options: ISPPullOptions;
  private spr: sprequest.ISPRequest;
  private agent: https.Agent;
  private utils: Utils;
  private apiSupportCheck: { [siteUrl: string]: boolean } = {};

  constructor(context: ISPPullContext, options: ISPPullOptions) {
    this.context = context;
    this.options = {
      ...options,
      dlRootFolder: options.dlRootFolder || '.downloads',
      metaFields: options.metaFields || []
    };
    this.utils = new Utils();
    this.agent = new https.Agent({
      rejectUnauthorized: false,
      keepAlive: true,
      keepAliveMsecs: 10000
    });
  }

  public async downloadFile(spFilePath: string, metadata?: IFileBasicMetadata): Promise<string> {
    this.spr = this.getCachedRequest();

    const spBaseFolderRegEx = new RegExp(decodeURIComponent('^' + this.options.spBaseFolder), 'gi');
    let spFilePathRelative = decodeURIComponent(spFilePath);
    if (['', '/'].indexOf(this.options.spBaseFolder) === -1) {
      spFilePathRelative = decodeURIComponent(spFilePath).replace(spBaseFolderRegEx, '');
    }

    let saveFilePath = path.join(this.options.dlRootFolder, spFilePathRelative);

    if (typeof this.options.omitFolderPath !== 'undefined') {
      // const omitFolderPath = path.resolve(this.options.omitFolderPath);
      saveFilePath = path.join(saveFilePath.replace(this.options.omitFolderPath, ''));
    }

    if (this.needToDownload(saveFilePath, metadata)) {
      const saveFolderPath = path.dirname(saveFilePath);
      await mkdirp(saveFolderPath);
      await this.download(spFilePath, saveFilePath);
    }

    return saveFilePath;
  }

  public async getFolderContent(spRootFolder: string): Promise<IContent> {
    let restUrl: string;
    this.spr = this.getCachedRequest();

    if (spRootFolder.charAt(spRootFolder.length - 1) === '/') {
      spRootFolder = spRootFolder.substring(0, spRootFolder.length - 1);
    }

    const folderInDocLibrary = await this.checkIfFolderInDocLibrary(spRootFolder).catch(() => false);
    const isModern = await this.checkModernApisSupport(this.context.siteUrl);

    if (folderInDocLibrary) {
      restUrl = this.utils.trimMultiline(`
        ${this.context.siteUrl}/_api/Web/${
          isModern
            ? 'GetFolderByServerRelativePath(DecodedUrl=@FolderServerRelativeUrl)'
            : 'GetFolderByServerRelativeUrl(@FolderServerRelativeUrl)'
        }
          ?$expand=Folders,Files,Folders/ListItemAllFields,Files/ListItemAllFields
          &$select=##MetadataSrt#
            Folders/ListItemAllFields/Id,
            Folders/Name,Folders/UniqueID,Folders/ID,Folders/ItemCount,Folders/ServerRelativeUrl,Folder/TimeCreated,Folder/TimeLastModified,
            Files/Name,Files/UniqueID,Files/ID,Files/ServerRelativeUrl,Files/Length,Files/TimeCreated,Files/TimeLastModified,Files/ModifiedBy
          &@FolderServerRelativeUrl='${this.utils.escapeURIComponent(spRootFolder)}'
      `);
      let metadataStr: string = this.options.metaFields.map(fieldName => {
        return 'Files/ListItemAllFields/' + fieldName;
      }).join(',');
      if (metadataStr.length > 0) {
        metadataStr += ',';
      }
      restUrl = restUrl.replace(/##MetadataSrt#/g, metadataStr);
    } else {
      restUrl = this.utils.trimMultiline(`
        ${this.context.siteUrl}/_api/Web/${
          isModern
            ? 'GetFolderByServerRelativePath(DecodedUrl=@FolderServerRelativeUrl)'
            : 'GetFolderByServerRelativeUrl(@FolderServerRelativeUrl)'
        }
          ?$expand=Folders,Files
          &$select=
            Folders/Name,Folders/UniqueID,Folders/ID,Folders/ItemCount,Folders/ServerRelativeUrl,Folder/TimeCreated,Folder/TimeLastModified,
            Files/Name,Files/UniqueID,Files/ID,Files/ServerRelativeUrl,Files/Length,Files/TimeCreated,Files/TimeLastModified,Files/ModifiedBy
          &@FolderServerRelativeUrl='${this.utils.escapeURIComponent(spRootFolder)}'
      `);
    }

    const response = await this.spr.get(restUrl, {
      agent: this.utils.isUrlHttps(restUrl) ? this.agent : undefined
    });

    const results: IContent = {
      folders: (response.body.d.Folders.results || []).filter((folder) => {
        if (folderInDocLibrary) {
          return typeof folder.ListItemAllFields.Id !== 'undefined';
        } else {
          return true;
        }
      }),
      files: (response.body.d.Files.results || []).map((file) => {
        if (folderInDocLibrary) {
          return {
            ...file,
            metadata: this.options.metaFields.reduce((meta, field) => {
              if (typeof file.ListItemAllFields !== 'undefined') {
                if (file.ListItemAllFields.hasOwnProperty(field)) {
                  meta[field] = file.ListItemAllFields[field];
                }
              }
              return meta;
            }, {})
          };
        } else {
          return {
            ...file,
            metafata: {}
          };
        }
      })
    };

    return results;

    // .catch((err) => {
    //   const message = err.message || err;
    //   console.log(colors.red.bold('\nError in getFolderContent:'), colors.red(message));
    // });
  }

  public async getContentWithCaml(): Promise<IContent> {
    this.spr = this.getCachedRequest();
    const digest = await this.spr.requestDigest(this.context.siteUrl);
    let restUrl = this.utils.trimMultiline(`
      ${this.context.siteUrl}/_api/Web/GetList(@DocLibUrl)/GetItems
        ?$select=##MetadataSrt#
          Name,UniqueID,ID,FileDirRef,FileRef,FSObjType,TimeCreated,TimeLastModified,Length,ModifiedBy,File/Length
        &$expand=Files/ListItemAllFields,File
        &@DocLibUrl='${this.utils.escapeURIComponent(this.options.spDocLibUrl)}'
    `);

    let metadataStr: string = this.options.metaFields.map((fieldName) => {
      // return `Files/ListItemAllFields/${fieldName}`;
      return `${fieldName}`;
    }).join(',');

    if (metadataStr.length > 0) {
      metadataStr += ',';
    }

    restUrl = restUrl.replace(/##MetadataSrt#/g, metadataStr);

    const response = await this.spr.post(restUrl, {
      body: {
        query: {
          __metadata: {
            type: 'SP.CamlQuery'
          },
          ViewXml: `<View Scope="Recursive"><Query><Where>${this.options.camlCondition}</Where></Query></View>`
        }
      },
      headers: {
        'X-RequestDigest': digest,
        'Accept': 'application/json; odata=verbose',
        'Content-Type': 'application/json; odata=verbose'
      },
      agent: this.utils.isUrlHttps(restUrl) ? this.agent : undefined
    });

    const spRootFolder = this.options.spRootFolder ? decodeURIComponent(this.options.spRootFolder) : undefined;

    const filesData: IFile[] = [];
    const foldersData: IFolder[] = [];
    response.body.d.results.forEach((item) => {
      // Exclude anything outside spRootFolder if provided
      if (spRootFolder && item.FileRef.indexOf(spRootFolder) !== 0) {
        return;
      }
      item.metadata = this.options.metaFields.reduce((meta, field) => {
        if (item.hasOwnProperty(field)) {
          meta[field] = item[field];
        }
        return meta;
      }, {});
      if (item.FSObjType === 0) {
        item.ServerRelativeUrl = item.FileRef;
        item.Length = item.File?.Length || '0';
        filesData.push(item);
      } else {
        foldersData.push(item);
      }
    });
    const results: IContent = {
      files: filesData,
      folders: foldersData
    };

    return results;

    // .catch((err) => {
    //   const message = err.message || err;
    //   console.log(colors.red.bold('\nError in getContentWithCaml:'), colors.red(message));
    // });
  }

  private async checkIfFolderInDocLibrary(spFolder: string): Promise<boolean> {
    this.spr = this.getCachedRequest();

    if (spFolder.charAt(spFolder.length - 1) === '/') {
      spFolder = spFolder.substring(0, spFolder.length - 1);
    }

    const isModern = await this.checkModernApisSupport(this.context.siteUrl);

    const restUrl = this.utils.trimMultiline(`
      ${this.context.siteUrl}/_api/Web/${
        isModern
          ? 'GetFolderByServerRelativePath(DecodedUrl=@FolderServerRelativeUrl)'
          : 'GetFolderByServerRelativeUrl(@FolderServerRelativeUrl)'
      }/listItemAllFields
        ?@FolderServerRelativeUrl='${this.utils.escapeURIComponent(spFolder)}'
    `);

    return new Promise((resolve) => {
      this.spr.get(restUrl, {
        agent: this.utils.isUrlHttps(restUrl) ? this.agent : undefined
      })
        .then(() => resolve(true))
        .catch(() => resolve(false));
    });
  }

  private async download(spFilePath: string, saveFilePath: string): Promise<string> {
    const isModern = await this.checkModernApisSupport(this.context.siteUrl);

    let restUrl: string = `${this.context.siteUrl}/_api/Web/` +
      `GetFileByServerRelativeUrl(@FileServerRelativeUrl)/$value` +
      `?@FileServerRelativeUrl='${this.utils.escapeURIComponent(spFilePath)}'`;

    if (isModern) {
      restUrl = `${this.context.siteUrl}/_api/Web/` +
        `GetFileByServerRelativePath(DecodedUrl=@FileServerRelativeUrl)/$value` +
        `?@FileServerRelativeUrl='${this.utils.escapeURIComponent(spFilePath)}'`;
    }

    let envProcessHeaders = {};
    try {
      // tslint:disable-next-line: no-string-literal
      envProcessHeaders = JSON.parse(process.env['_sp_request_headers'] || '{}');
    } catch (ex) { /**/ }

    return new Promise((resolve, reject) => {
      getAuth(this.context.siteUrl, this.context.creds)
        .then((auth) => {
          const options: request.OptionsWithUrl = {
            url: restUrl,
            method: 'GET',
            headers: {
              ...envProcessHeaders,
              ...auth.headers,
              'User-Agent': 'sppull'
            },
            encoding: null,
            strictSSL: false,
            gzip: true,
            agent: this.utils.isUrlHttps(this.context.siteUrl) ? this.agent : undefined,
            ...auth.options
          };
          request(options)
            .pipe(fs.createWriteStream(saveFilePath))
            .on('error', reject)
            .on('finish', () => resolve(saveFilePath));
        })
        .catch(reject);
    });
  }

  private needToDownload(saveFilePath: string, metadata?: IFileBasicMetadata): boolean {
    let stats: fs.Stats = null;
    let needDownload: boolean = true;

    if (typeof metadata !== 'undefined') {
      if (fs.existsSync(saveFilePath)) {
        stats = fs.statSync(saveFilePath);
        needDownload = false;
        if (typeof metadata.Length !== 'undefined') {
          if (stats.size !== parseInt(metadata.Length + '', 10)) {
            needDownload = true;
          }
        } else {
          needDownload = true;
        }
        if (typeof metadata.TimeLastModified !== 'undefined') {
          const timeLastModified = new Date(metadata.TimeLastModified);
          if (stats.mtime < timeLastModified) {
            needDownload = true;
          }
        }
      } else {
        needDownload = true;
      }
    }

    return needDownload;
  }

  private getCachedRequest() {
    return this.spr || sprequest.create(this.context.creds);
  }

  // checkModernApisSupport checks is `*ByServerRelativePath` methods exists on environment (were introduced in SPO, 2019)
  // to use these methods in preference to `*ByServerRelativeUrl` due to support for file and folder names with some special characters
  private async checkModernApisSupport(siteUrl: string): Promise<boolean> {
    const resultCache = this.apiSupportCheck[siteUrl.toLocaleLowerCase()];
    if (typeof resultCache !== 'undefined') {
      return resultCache;
    }

    const rootFolder = `/${siteUrl.split('/').slice(3).join('/')}`;
    const restUrl = `${siteUrl}/_api/Web/GetFolderByServerRelativePath(DecodedUrl=@FolderRelativeUrl)` +
      `?$select=Id&@FolderRelativeUrl='${this.utils.escapeURIComponent(rootFolder)}'`;

    this.spr = this.getCachedRequest();
    const result = await this.spr.get(restUrl, {
      agent: this.utils.isUrlHttps(restUrl) ? this.agent : undefined
    })
      .then((resp) => resp.statusCode === 200)
      .catch(() => false);

    this.apiSupportCheck[siteUrl.toLocaleLowerCase()] = result;

    return result;
  }

}
