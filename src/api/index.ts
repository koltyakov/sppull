import * as fs from 'fs';
import * as path from 'path';
import * as mkdirp from 'mkdirp';
import * as colors from 'colors';
import * as readline from 'readline';
import * as sprequest from 'sp-request';

import { ISPPullOptions, ISPPullContext, IFileBasicMetadata } from '../interfaces';

export default class RestAPI {

    private context: ISPPullContext;
    private options: ISPPullOptions;
    private spr: sprequest.ISPRequest;

    constructor(context: ISPPullContext, options: ISPPullOptions) {
        this.context = context;
        this.options = {
            ...options,
            dlRootFolder: options.dlRootFolder || '.downloads',
            metaFields: options.metaFields || []
        };
    }

    public downloadFile = (spFilePath: string, metadata?: IFileBasicMetadata): Promise<any> => {
        return new Promise((resolve, reject) => {
            this.spr = this.getCachedRequest();
            let restUrl: string = this.context.siteUrl +
                '/_api/Web/GetFileByServerRelativeUrl(@FileServerRelativeUrl)/OpenBinaryStream' +
                '?@FileServerRelativeUrl=\'' + encodeURIComponent(spFilePath) + '\'';

            let saveFilePath = path.join(
                this.options.dlRootFolder,
                decodeURIComponent(spFilePath).replace(decodeURIComponent(this.options.spBaseFolder), '')
            );

            if (typeof this.options.omitFolderPath !== 'undefined') {
                let omitFolderPath = path.resolve(this.options.omitFolderPath);
                saveFilePath = path.join(saveFilePath.replace(this.options.omitFolderPath, ''));
            }

            let stats: fs.Stats = null;
            let needDownload: boolean = true;

            if (typeof metadata !== 'undefined') {
                if (fs.existsSync(saveFilePath)) {
                    stats = fs.statSync(saveFilePath);
                    needDownload = false;
                    if (typeof metadata.Length !== 'undefined') {
                        // tslint:disable-next-line:radix
                        if (stats.size !== parseInt(metadata.Length + '')) {
                            needDownload = true;
                        }
                    } else {
                        needDownload = true;
                    }
                    if (typeof metadata.TimeLastModified !== 'undefined') {
                        let timeLastModified = new Date(metadata.TimeLastModified);
                        if (stats.mtime < timeLastModified) {
                            needDownload = true;
                        }
                    }
                } else {
                    needDownload = true;
                }
            }

            if (needDownload) {
                this.spr.get(restUrl, { encoding: null })
                    .then((response) => {
                        let saveFolderPath = path.dirname(saveFilePath);
                        if (/.json$/.test(saveFilePath)) {
                            response.body = JSON.stringify(response.body, null, 2);
                        }
                        if (/.map$/.test(saveFilePath)) {
                            response.body = JSON.stringify(response.body);
                        }
                        mkdirp(saveFolderPath, (err) => {
                            if (err) { reject(err); }
                            // tslint:disable-next-line:no-shadowed-variable
                            fs.writeFile(saveFilePath, response.body, (err) => {
                                if (err) { reject(err); }
                                resolve(saveFilePath);
                            });
                        });
                    })
                    .catch((err) => {
                        console.log(colors.red.bold('\nError in operations.downloadFile:'), colors.red(err.message));
                        reject(err.message);
                    });
            } else {
                resolve(saveFilePath);
            }
        });
    }

    public getFolderContent = (spRootFolder: string): Promise<any> => {
        return new Promise((resolve, reject) => {
            let restUrl: string;
            this.spr = this.getCachedRequest();

            if (spRootFolder.charAt(spRootFolder.length - 1) === '/') {
                spRootFolder = spRootFolder.substring(0, spRootFolder.length - 1);
            }

            restUrl = this.context.siteUrl + '/_api/Web/GetFolderByServerRelativeUrl(@FolderServerRelativeUrl)' +
                        '?$expand=Folders,Files,Folders/ListItemAllFields,Files/ListItemAllFields' +
                        '&$select=##MetadataSrt#' +
                        'Folders/Name,Folders/UniqueID,Folders/ID,Folders/ItemCount,Folders/ServerRelativeUrl,Folder/TimeCreated,Folder/TimeLastModified,' +
                        'Files/Name,Files/UniqueID,Files/ID,Files/ServerRelativeUrl,Files/Length,Files/TimeCreated,Files/TimeLastModified,Files/ModifiedBy' +
                        '&@FolderServerRelativeUrl=\'' + encodeURIComponent(spRootFolder) + '\'';

            let metadataStr: string = this.options.metaFields.map((fieldName) => {
                return 'Files/ListItemAllFields/' + fieldName;
            }).join(',');

            if (metadataStr.length > 0) {
                metadataStr += ',';
            }

            restUrl = restUrl.replace(/##MetadataSrt#/g, metadataStr);

            this.spr.get(restUrl)
                .then((response: any) => {
                    let results = {
                        folders: response.body.d.Folders.results || [],
                        files: (response.body.d.Files.results || []).map((file) => {
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

    public getContentWithCaml = (): Promise<any> => {
        return new Promise((resolve, reject) => {
            this.spr = this.getCachedRequest();
            this.spr.requestDigest(this.context.siteUrl)
                .then((digest) => {
                    let restUrl;
                    restUrl = this.context.siteUrl + '/_api/Web/GetList(@DocLibUrl)/GetItems' +
                            '?$select=' +
                                '##MetadataSrt#' +
                                'Name,UniqueID,ID,FileDirRef,FileRef,FSObjType,TimeCreated,TimeLastModified,Length,ModifiedBy' +
                            '&@DocLibUrl=\'' + encodeURIComponent(this.options.spDocLibUrl) + '\'';

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
                                ViewXml: '<View Scope=\'Recursive\'><Query><Where>' + this.options.camlCondition + '</Where></Query></View>'
                            }
                        },
                        headers: {
                            'X-RequestDigest': digest,
                            'accept': 'application/json; odata=verbose',
                            'content-type': 'application/json; odata=verbose'
                        }
                    });
                })
                .then((response) => {
                    let filesData = [];
                    let foldersData = [];
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
                    let results = {
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

    private getCachedRequest = (): sprequest.ISPRequest => {
        return this.spr || sprequest.create(this.context.creds);
    }

}
