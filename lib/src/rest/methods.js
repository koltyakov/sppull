var fs = require('fs');
var path = require('path');
var mkdirp = require('mkdirp');
var readline = require('readline');
var colors = require('colors');
var sprequest = require('sp-request');
var spr = null;
var ctxCached = null;

var spf = spf || {};
spf.rest = spf.rest || {};

spf.rest.RestApi = function() {
    var _that = this;
    var getCachedRequest = function(ctx) {
        var env = {};
        if (ctx.hasOwnProperty("domain")) {
            env.domain = ctx.domain;
        }
        if (ctx.hasOwnProperty("workstation")) {
            env.workstation = ctx.workstation;
        }
        if (JSON.stringify(ctxCached) === JSON.stringify(ctx)) {
            spr = spr || require("sp-request").create(ctx, env);
        } else {
            spr = require("sp-request").create(ctx, env);
        }
        return spr;
    };

    var operations = {};
    operations.downloadFile = function (context, options, spFilePath, callback) {

        var spFolderBase = options.spBaseFolder;
        var downloadRoot = options.dlRootFolder || "./downloads";

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
                console.log(colors.red.bold("\nError in operations.downloadFile:"), colors.red(err.message));
                if (callback && typeof callback === "function") {
                    callback(err.message);
                }
            });
    };
    operations.getFolderContent = function(context, options, spRootFolder, callback) {

        var metaFields = options.metaFields;

        var restUrl;
        spr = getCachedRequest(context);
        if (spRootFolder.charAt(spRootFolder.length - 1) === "/") {
            spRootFolder = spRootFolder.substring(0, spRootFolder.length - 1);
        }

        restUrl = context.siteUrl + "/_api/Web/GetFolderByServerRelativeUrl(@FolderServerRelativeUrl)" +
                    "?$expand=Folders,Files,Folders/ListItemAllFields,Files/ListItemAllFields" +
                    "&$select=" +
                    "##MetadataSrt#" +
                    "Folders/Name,Folders/UniqueID,Folders/ID,Folders/ItemCount,Folders/ServerRelativeUrl,Folder/TimeCreated,Folder/TimeModified," +
                    "Files/Name,Files/UniqueID,Files/ID,Files/ServerRelativeUrl,Files/Length,Files/TimeCreated,Files/TimeModified,Files/ModifiedBy" +
                    "&@FolderServerRelativeUrl='" + encodeURIComponent(spRootFolder) + "'";

        var metadataStr = "";
        if (metaFields.length > 0) {
            metadataStr = metaFields.map(function(fieldName) {
                return "Files/ListItemAllFields/" + fieldName;
            }).join(",") + ",";
        }
        restUrl = restUrl.replace(/##MetadataSrt#/g, metadataStr);

        spr.get(restUrl)
            .then(function (response) {

                var files = response.body.d.Files.results;
                if (metaFields.length > 0) {
                    files.forEach(function(data) {
                        data.metadata = {};
                        metaFields.forEach(function(fd) {
                            if (typeof data.ListItemAllFields !== "undefined") {
                                if (data.ListItemAllFields.hasOwnProperty(fd)) {
                                    data.metadata[fd] = data.ListItemAllFields[fd];
                                }
                            }
                        });
                    });
                }

                var results = {
                    folders: response.body.d.Folders.results,
                    files: files
                };

                if (callback && typeof callback === "function") {
                    callback(results);
                }
            })
            .catch(function (err) {
                console.log(colors.red.bold("\nError in operations.getFolderContent:"), colors.red(err.message));
                if (callback && typeof callback === "function") {
                    callback(err.message);
                }
            });
    };
    operations.getContentWithCaml = function(context, options, callback) {

        var spDocLibUrl = options.spDocLibUrl;
        var camlStr = options.camlCondition;
        var metaFields = options.metaFields;

        spr = getCachedRequest(context);
        spr.requestDigest(context.siteUrl)
            .then(function (digest) {
                var restUrl;
                restUrl = context.siteUrl + "/_api/Web/GetList(@DocLibUrl)/GetItems" +
                        "?$select=" +
                            "##MetadataSrt#" +
                            "Name,UniqueID,ID,FileDirRef,FileRef,FSObjType,TimeCreated,TimeModified,Length,ModifiedBy" +
                        "&@DocLibUrl='" + encodeURIComponent(spDocLibUrl) + "'";
                var metadataStr = "";
                if (metaFields.length > 0) {
                    metadataStr = metaFields.join(",") + ",";
                }
                restUrl = restUrl.replace(/##MetadataSrt#/g, metadataStr);
                return spr.post(restUrl, {
                    body: {
                        query: {
                            __metadata: {
                                type: "SP.CamlQuery"
                            },
                            ViewXml: "<View Scope='Recursive'><Query><Where>" + camlStr + "</Where></Query></View>" // <Eq><FieldRef Name='MyTestAttr'/><Value Type='Text'>MyTestValue</Value></Eq>
                        }
                    },
                    headers: {
                        "X-RequestDigest": digest,
                        "accept": "application/json; odata=verbose",
                        "content-type": "application/json; odata=verbose"
                    }
                });
            })
            .then(function (response) {
                var filesData = [];
                var foldersData = [];

                response.body.d.results.forEach(function(item) {
                    if (metaFields.length > 0) {
                        item.metadata = {};
                        metaFields.forEach(function(fd) {
                            if (item.hasOwnProperty(fd)) {
                                item.metadata[fd] = item[fd];
                            }
                        });
                    }
                    if (item.FSObjType === 0) {
                        item.ServerRelativeUrl = item.FileRef;
                        filesData.push(item);
                    } else {
                        foldersData.push(item);
                    }
                });

                var results = {
                    files: filesData,
                    folders: foldersData
                };

                if (callback && typeof callback === "function") {
                    callback(results);
                }
            })
            .catch(function (err) {
                console.log(colors.red.bold("\nError in operations.getContentWithCaml:"), colors.red(err.message));
                if (callback && typeof callback === "function") {
                    callback(err.message);
                }
            });
    };

    return operations;
};

module.exports = spf.rest;