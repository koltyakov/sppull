var fs = require('fs');
var path = require('path');
var mkdirp = require('mkdirp');
var colors = require('colors');
var readline = require('readline');
var Promise = require('bluebird');
var RestApi = require('./rest/methods').RestApi;

var sppull = function() {
    var _self = this;
    _self.restApi = new RestApi();

    // Queues >>>>
    var createFoldersQueue = function(foldersList, index, callback) {
        var spFolderPath = foldersList[index].ServerRelativeUrl;
        var spBaseFolder = _self.options.spBaseFolder;
        var downloadRoot = _self.options.dlRootFolder;
        var createFolder = function(spFolderPath, spBaseFolder, downloadRoot, callback) {
            var saveFolderPath = downloadRoot + "/" + decodeURIComponent(spFolderPath).replace(decodeURIComponent(spBaseFolder), "");
            mkdirp(saveFolderPath, function(err) {
                if (err) {
                    console.log(colors.red.bold("Error creating folder " + "`" + saveFolderPath + " `"), colors.red(err));
                    // throw err;
                };
                if (callback && typeof callback === "function") {
                    callback(saveFolderPath);
                }
            });
        };

        if (!_self.options.muteConsole) {
            readline.clearLine(process.stdout, 0);
            readline.cursorTo(process.stdout, 0, null);
            process.stdout.write(colors.green.bold("Creating folders: ") + (index + 1) + " out of " + foldersList.length);
        }

        createFolder(spFolderPath, spBaseFolder, downloadRoot, function(localFolderPath) {
            foldersList[index].SavedToLocalPath = localFolderPath;
            index += 1;
            if (index < foldersList.length) {
                createFoldersQueue(foldersList, index, callback);
            } else {

                if (!_self.options.muteConsole) {
                    process.stdout.write("\n");
                }

                if (callback && typeof callback === "function") {
                    callback(foldersList);
                }
            }
        });
    };
    var downloadFilesQueue = function(filesList, index, callback) {
        var spFilePath = filesList[index].ServerRelativeUrl;

        if (!_self.options.muteConsole) {
            readline.clearLine(process.stdout, 0);
            readline.cursorTo(process.stdout, 0, null);
            process.stdout.write(colors.green.bold("Downloading files: ") + (index + 1) + " out of " + filesList.length);
        }

        _self.restApi.downloadFile(_self.context, _self.options, spFilePath, function(localFilePath) {
            filesList[index].SavedToLocalPath = localFilePath;
            index += 1;
            if (index < filesList.length) {
                downloadFilesQueue(filesList, index, callback);
            } else {

                if (!_self.options.muteConsole) {
                    process.stdout.write("\n");
                }

                if (callback && typeof callback === "function") {
                    callback(filesList);
                }
            }
        });
    };
    var getStructureRecursive = function(foldersQueue, filesList, callback) {
        var exitQueue = true;
        var filesList = filesList || [];
        if (typeof _self.options.spRootFolder === "undefined" || _self.options.spRootFolder === "") {
            throw "The `spRootFolder` property should be provided in options.";
        }
        var spRootFolder;
        if (typeof foldersQueue === "undefined" || foldersQueue === null) {
            foldersQueue = [];
            spRootFolder = _self.options.spRootFolder;
            exitQueue = false;
        } else {
            foldersQueue.some(function(fi) {
                if (typeof fi.processed === "undefined") {
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
            var cntInQueue = 0;
            foldersQueue.forEach(function(folder) {
                if (folder.processed) {
                    cntInQueue += 1;
                }
            });

            if (!_self.options.muteConsole) {
                readline.clearLine(process.stdout, 0);
                readline.cursorTo(process.stdout, 0, null);
                process.stdout.write(colors.green.bold("Folders proceeding: ") + cntInQueue + " out of " + foldersQueue.length + colors.gray(" [reqursive scanning...]"));
                // " | files found: " + filesList.length +
            }

            _self.restApi.getFolderContent(_self.context, _self.options, spRootFolder, function(results) {
                (results.folders || []).forEach(function(folder) {
                    var folderElement  = {
                        folder: folder,
                        serverRelativeUrl: folder.ServerRelativeUrl,
                        processed: false
                    };
                    foldersQueue.push(folderElement);
                });
                filesList = filesList.concat(results.files || []);
                getStructureRecursive(foldersQueue, filesList, callback);
            });

        } else {

            if (!_self.options.muteConsole) {
                process.stdout.write("\n");
            }

            if (callback && typeof callback === "function") {
                var foldersList = foldersQueue.map(function(folder) {
                    return folder.folder;
                });
                callback({
                    files: filesList,
                    folders: foldersList
                });
            }
        }
    };
    // <<<< Queues

    // Runners >>>>
    var runCreateFoldersRecursively = function() {
        return new Promise(function(resolve, reject) {
            getStructureRecursive(null, null, function(data) {
                if ((data.folders || []).length > 0) {
                    createFoldersQueue(data.folders, 0, resolve);
                } else {
                    resolve([]);
                }
            });
        });
    };
    var runDownloadFilesRecursively = function() {
        return new Promise(function(resolve, reject) {
            getStructureRecursive(null, null, function(data) {
                var downloadMyFilesHandler = function() {
                    if ((data.files || []).length > 0) {
                        downloadFilesQueue(data.files, 0, resolve);
                    } else {
                        resolve([]);
                    }
                };
                if (_self.options.createEmptyFolders) {
                    if ((data.folders || []).length > 0) {
                        createFoldersQueue(data.folders, 0, function() {
                            downloadMyFilesHandler();
                        });
                    } else {
                        downloadMyFilesHandler();
                    }
                } else {
                    downloadMyFilesHandler();
                }
            });
        });
    };
    var runDownloadFilesFlat = function() {
        return new Promise(function(resolve, reject) {
            _self.restApi.getFolderContent(_self.context, _self.options, function(data) {
                var downloadMyFilesHandler = function() {
                    if ((data.files || []).length > 0) {
                        downloadFilesQueue(data.files, 0, resolve);
                    } else {
                        resolve([]);
                    }
                };
                if (_self.options.createEmptyFolders) {
                    if ((data.folders || []).length > 0) {
                        createFoldersQueue(data.folders, 0, function() {
                            downloadMyFilesHandler();
                        });
                    } else {
                        downloadMyFilesHandler();
                    }
                } else {
                    downloadMyFilesHandler();
                }
            });
        });
    };
    var runDownloadStrictObjects = function() {
        return new Promise(function(resolve, reject) {
            var filesList = _self.options.strictObjects.filter(function(d) {
                var pathArr = d.split("/");
                return pathArr[pathArr.length - 1].indexOf(".") !== -1;
            }).map(function(d) {
                return {
                    ServerRelativeUrl: d
                };
            });
            if (filesList.length > 0) {
                downloadFilesQueue(filesList, 0, resolve);
            } else {
                resolve([]);
            }
        });
    };
    var runDownloadCamlObjects = function() {
        return new Promise(function(resolve, reject) {
            _self.restApi.getContentWithCaml(_self.context, _self.options, function(data) {
                var downloadMyFilesHandler = function() {
                    if ((data.files || []).length > 0) {
                        downloadFilesQueue(data.files, 0, resolve);
                    } else {
                        resolve([]);
                    }
                };
                if (_self.options.createEmptyFolders) {
                    if ((data.folders || []).length > 0) {
                        createFoldersQueue(data.folders, 0, function() {
                            downloadMyFilesHandler();
                        });
                    } else {
                        downloadMyFilesHandler();
                    }
                } else {
                    downloadMyFilesHandler();
                }
            });
        });
    };
    // <<<< Runners

    _self.sppull = function(context, options) {

        _self.context = context;
        _self.options = options;

        _self.options.spHostName = _self.context.siteUrl
            .replace("http://", "")
            .replace("https://", "")
            .split("/")[0];
        _self.options.spRelativeBase = _self.context.siteUrl
            .replace("http://", "")
            .replace("https://", "")
            .replace(_self.options.spHostName, "");

        if (_self.options.spRootFolder) {
            if (_self.options.spRootFolder.indexOf(_self.options.spRelativeBase) !== 0) {
                _self.options.spRootFolder = (_self.options.spRelativeBase + "/" + _self.options.spRootFolder).replace(/\/\//g, "/");
            } else {
                if (_self.options.spRootFolder.charAt(0) !== "/") {
                    _self.options.spRootFolder = "/" + _self.options.spRootFolder;
                }
            }
        }

        if (_self.options.spBaseFolder) {
            if (_self.options.spBaseFolder.indexOf(_self.options.spRelativeBase) !== 0) {
                _self.options.spBaseFolder = (_self.options.spRelativeBase + "/" + _self.options.spBaseFolder).replace(/\/\//g, "/");
            }
        } else {
            _self.options.spBaseFolder = _self.options.spRootFolder;
        }

        _self.options.dlRootFolder = _self.options.dlRootFolder || "./Downloads";

        if (typeof _self.options.recursive === "undefined") {
            _self.options.recursive = true;
        }

        if (typeof _self.options.foderStructureOnly === "undefined") {
            _self.options.foderStructureOnly = false;
        }

        if (typeof _self.options.createEmptyFolders === "undefined") {
            _self.options.createEmptyFolders = true;
        }

        if (typeof _self.options.metaFields === "undefined") {
            _self.options.metaFields = [];
        }

        if (typeof _self.options.restCondition === "undefined") {
            _self.options.restCondition = "";
        }

        if (typeof _self.options.muteConsole === "undefined") {
            _self.options.muteConsole = false;
        }

        // ====

        if (typeof _self.options.camlCondition !== "undefined" && _self.options.camlCondition !== "" && typeof _self.options.spDocLibUrl !== "undefined" && _self.options.spDocLibUrl !== "") {

            return runDownloadCamlObjects();

        } else {

            if (typeof _self.options.strictObjects !== "undefined" && Array.isArray([_self.options.strictObjects])) {

                _self.options.strictObjects.forEach(function(strictObject, i) {
                    if (typeof strictObject === "string") {
                        if (strictObject.indexOf(_self.options.spRootFolder) !== 0) {
                            strictObject = (_self.options.spRootFolder + "/" + strictObject).replace(/\/\//g, "/");
                        }
                        _self.options.strictObjects[i] = strictObject;
                        // console.log(_self.options.spRootFolder, _self.options.strictObjects[i]);
                    }
                });

                return runDownloadStrictObjects();
            } else {
                if (!_self.options.foderStructureOnly) {
                    if (_self.options.recursive) {
                        return runDownloadFilesRecursively();
                    } else {
                        return runDownloadFilesFlat();
                    }
                } else {
                    return runCreateFoldersRecursively();
                }
            }

        }
    };

    return _self.sppull;
};

module.exports.sppull = new sppull();