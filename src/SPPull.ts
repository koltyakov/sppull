import * as fs from 'fs';
import * as path from 'path';
import * as mkdirp from 'mkdirp';
import * as colors from 'colors';
import * as readline from 'readline';

import RestAPI from './api';
import { ISPPullOptions, ISPPullContext } from './interfaces';
import { IFileBasicMetadata } from './interfaces';

export class Download {

    private context: ISPPullContext;
    private options: ISPPullOptions;
    private restApi: RestAPI;

    public sppull = (context: ISPPullContext, options: ISPPullOptions): Promise<any> => {

        this.context = context;
        this.options = options;

        if (typeof this.context.creds === 'undefined') {
            this.context.creds = {
                ...(this.context as any)
            };
        }

        this.options.spHostName = this.context.siteUrl
            .replace('http://', '')
            .replace('https://', '')
            .split('/')[0];

        this.options.spRelativeBase = this.context.siteUrl
            .replace('http://', '')
            .replace('https://', '')
            .replace(this.options.spHostName, '');

        if (this.options.spRootFolder) {
            if (this.options.spRootFolder.indexOf(this.options.spRelativeBase) !== 0) {
                this.options.spRootFolder = (this.options.spRelativeBase + '/' + this.options.spRootFolder).replace(/\/\//g, '/');
            } else {
                if (this.options.spRootFolder.charAt(0) !== '/') {
                    this.options.spRootFolder = '/' + this.options.spRootFolder;
                }
            }
        }

        if (this.options.spBaseFolder) {
            if (this.options.spBaseFolder.indexOf(this.options.spRelativeBase) !== 0) {
                this.options.spBaseFolder = (this.options.spRelativeBase + '/' + this.options.spBaseFolder).replace(/\/\//g, '/');
            }
        } else {
            this.options.spBaseFolder = this.options.spRootFolder;
        }

        this.options.dlRootFolder = this.options.dlRootFolder || './Downloads';

        if (typeof this.options.recursive === 'undefined') {
            this.options.recursive = true;
        }

        if (typeof this.options.foderStructureOnly === 'undefined') {
            this.options.foderStructureOnly = false;
        }

        if (typeof this.options.createEmptyFolders === 'undefined') {
            this.options.createEmptyFolders = true;
        }

        if (typeof this.options.metaFields === 'undefined') {
            this.options.metaFields = [];
        }

        if (typeof this.options.restCondition === 'undefined') {
            this.options.restCondition = '';
        }

        if (this.options.spDocLibUrl) {
            if (this.options.spDocLibUrl.indexOf(this.options.spRelativeBase) !== 0) {
                this.options.spDocLibUrl = (this.options.spRelativeBase + '/' + this.options.spDocLibUrl).replace(/\/\//g, '/');
            } else {
                if (this.options.spDocLibUrl.charAt(0) !== '/') {
                    this.options.spDocLibUrl = '/' + this.options.spDocLibUrl;
                }
            }
            this.options.spDocLibUrl = encodeURIComponent(this.options.spDocLibUrl);
        }

        if (typeof this.options.muteConsole === 'undefined') {
            this.options.muteConsole = false;
        }

        // Defaults <<<

        this.restApi = new RestAPI(this.context, this.options);

        if (
            typeof this.options.camlCondition !== 'undefined' && this.options.camlCondition !== '' &&
            typeof this.options.spDocLibUrl !== 'undefined' && this.options.spDocLibUrl !== ''
        ) {
            return this.runDownloadCamlObjects();
        } else {
            if (typeof this.options.strictObjects !== 'undefined' && Array.isArray([this.options.strictObjects])) {
                this.options.strictObjects.forEach((strictObject, i) => {
                    if (typeof strictObject === 'string') {
                        if (strictObject.indexOf(this.options.spRootFolder) !== 0) {
                            strictObject = (this.options.spRootFolder + '/' + strictObject).replace(/\/\//g, '/');
                        }
                        this.options.strictObjects[i] = strictObject;
                    }
                });
                return this.runDownloadStrictObjects();
            } else {
                if (!this.options.foderStructureOnly) {
                    if (this.options.recursive) {
                        return this.runDownloadFilesRecursively();
                    } else {
                        return this.runDownloadFilesFlat();
                    }
                } else {
                    return this.runCreateFoldersRecursively();
                }
            }

        }
    }

    private createFolder = (spFolderPath: string, spBaseFolder: string, downloadRoot: string): Promise<any> => {
        return new Promise((resolve, reject) => {
            let saveFolderPath: string = downloadRoot + '/' + decodeURIComponent(spFolderPath).replace(decodeURIComponent(spBaseFolder), '');
            mkdirp(saveFolderPath, (err) => {
                if (err) {
                    console.log(colors.red.bold('Error creating folder ' + '`' + saveFolderPath + ' `'), colors.red(err));
                    reject(err);
                }
                resolve(saveFolderPath);
            });
        });
    }

    // Queues >>>>
    private createFoldersQueue = (foldersList: any[], index: number = 0): Promise<any> => {
        return new Promise((resolve, reject) => {
            let spFolderPath = foldersList[index].ServerRelativeUrl;
            let spBaseFolder = this.options.spBaseFolder;
            let downloadRoot = this.options.dlRootFolder;

            if (!this.options.muteConsole) {
                readline.clearLine(process.stdout, 0);
                readline.cursorTo(process.stdout, 0, null);
                process.stdout.write(colors.green.bold('Creating folders: ') + (index + 1) + ' out of ' + foldersList.length);
            }

            this.createFolder(spFolderPath, spBaseFolder, downloadRoot)
                .then((localFolderPath: string) => {
                    foldersList[index].SavedToLocalPath = localFolderPath;
                    index += 1;
                    if (index < foldersList.length) {
                        resolve(this.createFoldersQueue(foldersList, index));
                    } else {
                        if (!this.options.muteConsole) {
                            process.stdout.write('\n');
                        }
                        resolve(foldersList);
                    }
                });
        });
    }

    private downloadFilesQueue = (filesList: IFileBasicMetadata[], index: number = 0): Promise<any> => {
        return new Promise((resolve, reject) => {
            let spFilePath = filesList[index].ServerRelativeUrl;
            if (!this.options.muteConsole) {
                readline.clearLine(process.stdout, 0);
                readline.cursorTo(process.stdout, 0, null);
                process.stdout.write(colors.green.bold('Downloading files: ') + (index + 1) + ' out of ' + filesList.length);
            }
            this.restApi.downloadFile(spFilePath, filesList[index])
                .then((localFilePath) => {
                    filesList[index].SavedToLocalPath = localFilePath;
                    index += 1;
                    if (index < filesList.length) {
                        resolve(this.downloadFilesQueue(filesList, index));
                    } else {
                        if (!this.options.muteConsole) {
                            process.stdout.write('\n');
                        }
                        resolve(filesList);
                    }
                });
        });
    }

    private getStructureRecursive = (root: boolean = true, foldersQueue: any[] = [], filesList: any[] = []): Promise<any> => {
        return new Promise((resolve, reject) => {
            let exitQueue = true;
            if (typeof this.options.spRootFolder === 'undefined' || this.options.spRootFolder === '') {
                reject('The `spRootFolder` property should be provided in options.');
            }
            let spRootFolder;

            if (foldersQueue.length === 0) {
                spRootFolder = this.options.spRootFolder;
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

                if (!this.options.muteConsole) {
                    readline.clearLine(process.stdout, 0);
                    readline.cursorTo(process.stdout, 0, null);
                    process.stdout.write(colors.green.bold('Folders proceeding: ') + cntInQueue + ' out of ' + foldersQueue.length + colors.gray(' [recursive scanning...]'));
                }

                this.restApi.getFolderContent(spRootFolder)
                    .then((results) => {
                        (results.folders || []).forEach((folder) => {
                            let folderElement = {
                                folder: folder,
                                serverRelativeUrl: folder.ServerRelativeUrl,
                                processed: false
                            };
                            foldersQueue.push(folderElement);
                        });
                        filesList = filesList.concat(results.files || []);
                        resolve(this.getStructureRecursive(false, foldersQueue, filesList));
                    });

            } else {
                if (!this.options.muteConsole) {
                    process.stdout.write('\n');
                }
                let foldersList = foldersQueue.map((folder) => {
                    return folder.folder;
                });
                resolve({
                    files: filesList,
                    folders: foldersList
                });
            }
        });
    }

    // <<<< Queues

    // Runners >>>>
    private runCreateFoldersRecursively = (): Promise<any> => {
        return new Promise((resolve, reject) => {
            return this.getStructureRecursive()
                .then((data) => {
                    if ((data.folders || []).length > 0) {
                        resolve(this.createFoldersQueue(data.folders, 0));
                    } else {
                        resolve([]);
                    }
                });
        });
    }

    private downloadMyFilesHandler = (data): Promise<any> => {
        return new Promise((resolve, reject) => {
            if ((data.files || []).length > 0) {
                resolve(this.downloadFilesQueue(data.files, 0));
            } else {
                resolve([]);
            }
        });
    }

    private runDownloadFilesRecursively(): Promise<any> {
        return this.getStructureRecursive()
            .then((data) => {
                if (this.options.createEmptyFolders) {
                    if ((data.folders || []).length > 0) {
                        return this.createFoldersQueue(data.folders, 0)
                            .then(() => {
                                return this.downloadMyFilesHandler(data);
                            });
                    } else {
                        return this.downloadMyFilesHandler(data);
                    }
                } else {
                    return this.downloadMyFilesHandler(data);
                }
            });
    }

    private runDownloadFilesFlat = (): Promise<any> => {
        if (typeof this.options.spRootFolder === 'undefined' || this.options.spRootFolder === '') {
            throw 'The `spRootFolder` property should be provided in options.';
        }
        return this.restApi.getFolderContent(this.options.spRootFolder)
            .then((data) => {
                if (this.options.createEmptyFolders) {
                    if ((data.folders || []).length > 0) {
                        return this.createFoldersQueue(data.folders, 0)
                            .then(() => {
                                return this.downloadMyFilesHandler(data);
                            });
                    } else {
                        return this.downloadMyFilesHandler(data);
                    }
                } else {
                    return this.downloadMyFilesHandler(data);
                }
            });
    }

    private runDownloadStrictObjects = (): Promise<any> => {
        return new Promise((resolve, reject) => {
            let filesList: IFileBasicMetadata[] = this.options.strictObjects.filter((d) => {
                let pathArr = d.split('/');
                return pathArr[pathArr.length - 1].indexOf('.') !== -1;
            }).map((d) => {
                return {
                    ServerRelativeUrl: d
                };
            });
            if (filesList.length > 0) {
                resolve(this.downloadFilesQueue(filesList, 0));
            } else {
                resolve([]);
            }
        });
    }

    private runDownloadCamlObjects = (): Promise<any> => {
        return this.restApi.getContentWithCaml()
            .then((data) => {
                if (this.options.createEmptyFolders) {
                    if ((data.folders || []).length > 0) {
                        return this.createFoldersQueue(data.folders, 0)
                            .then(() => {
                                return this.downloadMyFilesHandler(data);
                            });
                    } else {
                        return this.downloadMyFilesHandler(data);
                    }
                } else {
                    return this.downloadMyFilesHandler(data);
                }
            });
    }

    // <<<< Runners

}

export { ISPPullOptions, ISPPullContext } from './interfaces';