var fs = require("fs");
var path = require("path");
var URL = require('url-parse');
var mkdirp = require('mkdirp');
var Promise = require("bluebird");
var sprequest = require('sp-request');
var spr = null;
var ctxCached = null;

Promise.longStackTraces();

var sppull = function() {
    var _self = this;

    var RestOperations = function() {
        var _that = this;
        var getCachedRequest = function(ctx) {
            if (JSON.stringify(ctxCached) === JSON.stringify(ctx)) {
                spr = spr || require("sp-request").create(ctx);
            } else {
                spr = require("sp-request").create(ctx);
            }
            return spr;
        };

        var operations = {};
        operations.downloadFile = function (context, spFilePath, spFolderBase, downloadRoot, callback) {
            var downloadRoot = downloadRoot || "./Downloads";
            var restUrl;
            spr = getCachedRequest(context);
            restUrl = context.siteUrl + "/_api/Web/GetFileByServerRelativeUrl(@FileServerRelativeUrl)/OpenBinaryStream" +
                                        "?@FileServerRelativeUrl='" + encodeURIComponent(spFilePath) + "'";
            spr.get(restUrl, { encoding: null })
                .then(function (response) {
                    var saveFilePath = downloadRoot + "/" + decodeURIComponent(spFilePath).replace(decodeURIComponent(spFolderBase), "");
                    var saveFolderPath = path.dirname(saveFilePath);
                    mkdirp(saveFolderPath, function(err) {
                        fs.writeFile(saveFilePath, response.body, function(err) {
                            if (err) {
                                throw err;
                            };
                            if (callback && typeof callback === "function") {
                                callback(saveFilePath);
                            }
                        });
                    });
                })
                .catch(function (err) {
                    console.log("Error in operations.downloadFile:", err.message);
                    if (callback && typeof callback === "function") {
                        callback(err.message);
                    }
                });
        };
        operations.getFolderContent = function(context, spRootFolder, callback) {
            var restUrl;
            spr = getCachedRequest(context);
            if (spRootFolder.charAt(spRootFolder.length - 1) === "/") {
                spRootFolder = spRootFolder.substring(0, spRootFolder.length - 1);
            }

            restUrl = context.siteUrl + "/_api/Web/GetFolderByServerRelativeUrl(@FolderServerRelativeUrl)" + 
                      "?$expand=Folders,Files,Folders/ListItemAllFields,Files/ListItemAllFields" + // ,Folders/Folders,Folders/Files, Folders/ListItemAllFields,Files/ListItemAllFields,Folders/Properties,Files/Properties
                      "&$select=" + 
                           "Folders/Name,Folders/UniqueID,Folders/ID,Folders/ItemCount,Folders/ServerRelativeUrl,Folder/TimeCreated,Folder/TimeModified," +
                           "Files/Name,Files/UniqueID,Files/ID,Files/ServerRelativeUrl,Files/Length,Files/TimeCreated,Files/TimeModified,Files/ModifiedBy" +
                      "&@FolderServerRelativeUrl='" + encodeURIComponent(spRootFolder) + "'";

            spr.get(restUrl)
                .then(function (response) {
                    var results = {
                        folders: response.body.d.Folders.results,
                        files: response.body.d.Files.results
                    };

                    if (callback && typeof callback === "function") {
                        callback(results);
                    }
                })
                .catch(function (err) {
                    console.log("Error in operations.getFolderContent:", err.message);
                    if (callback && typeof callback === "function") {
                        callback(err.message);
                    }
                });
        };
        return operations;
    };
    var restOperations = new RestOperations();

    var createFoldersQueue = function(foldersList, index, callback) {
        var spFolderPath = foldersList[index].ServerRelativeUrl;
        var spBaseFolder = _self.options.spBaseFolder;
        var downloadRoot = _self.options.dlRootFolder; 
        var createFolder = function(spFolderPath, spBaseFolder, downloadRoot, callback) {
            var saveFolderPath = downloadRoot + "/" + decodeURIComponent(spFolderPath).replace(decodeURIComponent(spBaseFolder), "");
            var saveFolderPath = path.dirname(saveFolderPath);
            mkdirp(saveFolderPath, function(err) {
                if (err) {
                    console.log("Error creating folder " + "`" + saveFolderPath + " `", err);
                    // throw err;
                };
                if (callback && typeof callback === "function") {
                    callback(saveFolderPath);
                }
            });
        };
        createFolder(spFolderPath, spBaseFolder, downloadRoot, function(localFolderPath) {
            foldersList[index].SavedToLocalPath = localFolderPath;
            index += 1;
            if (index < foldersList.length) {
                createFoldersQueue(foldersList, index, callback);
            } else {
                if (callback && typeof callback === "function") {
                    callback(foldersList);
                }
            }
        });
    };
    var downloadFilesQueue = function(filesList, index, callback) {
        var spFilePath = filesList[index].ServerRelativeUrl;
        var spBaseFolder = _self.options.spBaseFolder;
        var downloadRoot = _self.options.dlRootFolder; 
        restOperations.downloadFile(_self.context, spFilePath, spBaseFolder, downloadRoot, function(localFilePath) {
            filesList[index].SavedToLocalPath = localFilePath;
            index += 1;
            if (index < filesList.length) {
                downloadFilesQueue(filesList, index, callback);
            } else {
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
            spRootFolder = _self.options.spRootFolder
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
            restOperations.getFolderContent(_self.context, spRootFolder, function(results) {
                (results.folders || []).forEach(function(folder) {
                    var folderElement  = {
                        folder: folder,
                        serverRelativeUrl: folder.ServerRelativeUrl
                    };
                    folderElement.processed = (folder.ItemCount === 0);
                    foldersQueue.push(folderElement);
                });
                filesList = filesList.concat(results.files || []);
                getStructureRecursive(foldersQueue, filesList, callback);
            });
        } else {
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
                if ((data.files || []).length > 0) {
                    downloadFilesQueue(data.files, 0, resolve);
                } else {
                    resolve([]);
                }
            });
        });
    };
    var runDownloadFilesFlat = function() {
        return new Promise(function(resolve, reject) {
            restOperations.getFolderContent(_self.context, _self.options.spRootFolder, function(data) {
                if ((data.files || []).length > 0) {
                    downloadFilesQueue(data.files, 0, resolve);
                } else {
                    resolve([]);
                }
            });
        });
    };    

    _self.sppull = function(context, options) {

        _self.context = context;
        _self.options = options;


        _self.options.dlRootFolder = _self.options.dlRootFolder || "./Downloads";

        _self.options.spBaseFolder = _self.options.spBaseFolder || _self.options.spRootFolder;
        if (typeof _self.options.recursive === "undefined") {
            _self.options.recursive = true;
        }

        if (typeof _self.options.foderStructureOnly === "undefined") {
            _self.options.foderStructureOnly = false;
        }

        if (!_self.options.foderStructureOnly) {
            if (_self.options.recursive) {
                return runDownloadFilesRecursively();
            } else {
                return runDownloadFilesFlat();
            }
        } else {
            return runCreateFoldersRecursively();
        }
    };

    return _self.sppull;
};

module.exports.sppull = new sppull();