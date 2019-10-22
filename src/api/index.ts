import * as fs from 'fs';
import * as path from 'path';
import * as https from 'https';
import * as mkdirp from 'mkdirp';
import * as colors from 'colors';
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

  public downloadFile(spFilePath: string, metadata?: IFileBasicMetadata): Promise<string> {
    return new Promise((resolve, reject) => {
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

        mkdirp(saveFolderPath, (err) => {
          if (err) {
            console.log(colors.red.bold('\nError in operations.downloadFile:'), colors.red(err));
            return reject(err);
          }

          // ToDo: Check the most effective approach
          // vs:
          //  - memory consumptions
          //  - speed

          // If a file is greater than 20 MB than use streams
          // tslint:disable-next-line:radix
          const filesize: number = parseInt(metadata.Length + '');
          if (filesize > 20000000) {

            // console.log('Download using streaming');

            // Download using streaming
            this.downloadAsStream(spFilePath, saveFilePath)
              .then(() => resolve(saveFilePath))
              .catch((error) => {
                console.log(colors.red.bold('\nError in operations.downloadFile:'), colors.red(error.message));
                reject(error);
              });

          } else {

            // console.log('Download simple');

            // Download using sp-request, without streaming, consumes lots of memory in case of large files
            this.downloadSimple(spFilePath, saveFilePath)
              .then(() => resolve(saveFilePath))
              .catch((error) => {
                console.log(colors.red.bold('\nError in operations.downloadFile:'), colors.red(error.message));
                reject(error);
              });

          }

        });

      } else {
        resolve(saveFilePath);
      }
    });
  }

  public async getFolderContent(spRootFolder: string): Promise<IContent> {
    let restUrl: string;
    this.spr = this.getCachedRequest();

    if (spRootFolder.charAt(spRootFolder.length - 1) === '/') {
      spRootFolder = spRootFolder.substring(0, spRootFolder.length - 1);
    }

    const folderInDocLibrary = await this.checkIfFolderInDocLibrary(spRootFolder).catch(() => false);

    if (folderInDocLibrary) {
      restUrl = this.utils.trimMultiline(`
        ${this.context.siteUrl}/_api/Web/GetFolderByServerRelativeUrl(@FolderServerRelativeUrl)
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
        ${this.context.siteUrl}/_api/Web/GetFolderByServerRelativeUrl(@FolderServerRelativeUrl)
          ?$expand=Folders,Files
          &$select=
            Folders/Name,Folders/UniqueID,Folders/ID,Folders/ItemCount,Folders/ServerRelativeUrl,Folder/TimeCreated,Folder/TimeLastModified,
            Files/Name,Files/UniqueID,Files/ID,Files/ServerRelativeUrl,Files/Length,Files/TimeCreated,Files/TimeLastModified,Files/ModifiedBy
          &@FolderServerRelativeUrl='${this.utils.escapeURIComponent(spRootFolder)}'
      `);
    }

    return new Promise((resolve, reject) => {
      this.spr.get(restUrl, {
        agent: this.utils.isUrlHttps(restUrl) ? this.agent : undefined
      })
        .then((response) => {
          const results = {
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
          resolve(results);
        })
        .catch((err) => {
          console.log(colors.red.bold('\nError in getFolderContent:'), colors.red(err.message));
          reject(err.message);
        });
    });
  }

  public getContentWithCaml(): Promise<IContent> {
    return new Promise((resolve, reject) => {
      this.spr = this.getCachedRequest();
      this.spr.requestDigest(this.context.siteUrl)
        .then((digest) => {
          let restUrl = this.utils.trimMultiline(`
            ${this.context.siteUrl}/_api/Web/GetList(@DocLibUrl)/GetItems
              ?$select=##MetadataSrt#
                Name,UniqueID,ID,FileDirRef,FileRef,FSObjType,TimeCreated,TimeLastModified,Length,ModifiedBy
              &@DocLibUrl='${this.utils.escapeURIComponent(this.options.spDocLibUrl)}'
          `);

          let metadataStr: string = this.options.metaFields.map((fieldName) => {
            return `Files/ListItemAllFields/${fieldName}`;
          }).join(',');

          if (metadataStr.length > 0) {
            metadataStr += ',';
          }

          restUrl = restUrl.replace(/##MetadataSrt#/g, metadataStr);

          return this.spr.post(restUrl, {
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
        })
        .then((response) => {
          const filesData: IFile[] = [];
          const foldersData: IFolder[] = [];
          response.body.d.results.forEach((item) => {
            item.metadata = this.options.metaFields.reduce((meta, field) => {
              if (item.hasOwnProperty(field)) {
                meta[field] = item[field];
              }
              return meta;
            }, {});
            if (item.FSObjType === 0) {
              item.ServerRelativeUrl = item.FileRef;
              filesData.push(item);
            } else {
              foldersData.push(item);
            }
          });
          const results: IContent = {
            files: filesData,
            folders: foldersData
          };
          resolve(results);
        })
        .catch((err) => {
          console.log(colors.red.bold('\nError in getContentWithCaml:'), colors.red(err.message));
          reject(err.message);
        });
    });
  }

  private checkIfFolderInDocLibrary(spFolder: string): Promise<boolean> {
    this.spr = this.getCachedRequest();

    if (spFolder.charAt(spFolder.length - 1) === '/') {
      spFolder = spFolder.substring(0, spFolder.length - 1);
    }

    const restUrl = this.utils.trimMultiline(`
      ${this.context.siteUrl}/_api/Web/GetFolderByServerRelativeUrl(@FolderServerRelativeUrl)/listItemAllFields
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

  private downloadAsStream(spFilePath: string, saveFilePath: string): Promise<string> {
    const restUrl: string = `${this.context.siteUrl}/_api/Web/` +
      `GetFileByServerRelativeUrl(@FileServerRelativeUrl)/$value` +
      `?@FileServerRelativeUrl='${this.utils.escapeURIComponent(spFilePath)}'`;

    let envProcessHeaders = {};
    try {
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

  private downloadSimple(spFilePath: string, saveFilePath: string): Promise<string> {
    const restUrl: string = `${this.context.siteUrl}/_api/Web/` +
      `GetFileByServerRelativeUrl(@FileServerRelativeUrl)/OpenBinaryStream` +
      `?@FileServerRelativeUrl='${this.utils.escapeURIComponent(spFilePath)}'`;

    let envProcessHeaders = {};
    try {
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

}
