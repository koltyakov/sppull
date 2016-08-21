# SPPull - simple client to pull and download files from SharePoint

Node.js module for downloading files from SharePoint document libraries.

## Supported SharePoint versions:
- SharePoint Online
- SharePoint 2013
- SharePoint 2016

## How to use:

### Install:
```bash
npm install sppull --save-dev
```
### Usage:
```javascript
var sppull = require("sppull").sppull;

sppull(context, options)
    .then(successHandler, errorHandler);
 
 /* OR */
sppull(context, options)
    .then(successHandler)
    .catch(errorHandler);
```
#### Arguments:

##### Context:
- `siteUrl` - SharePoint site (SPWeb) url, string, required
- `username` - user name for SP authentication, string, required
- `password` - password, string, required

##### Options:
- `spRootFolder` - root folder in SharePoint to pull from, string, required
- `dlRootFolder` - local root folder where files and folders will be saved to, string, required
- `spBaseFolder` - base folder path which is omitted then saving files locally, string, optional
- `recursive` - to pull all files and folders recursively, boolean, optional, default is `true`
- `ignoreEmptyFolders` - to ignore local creation of SharePoint empty folders, optional, default is `true`
- `foderStructureOnly` - to ignore files, recreate only folders' structure, optional, default is `false`
- `strictObjects` - array of files and folders relative paths within the `spRootFolder`, optional, array of strings
- `restCondition` - SharePoint REST filter conditions to use, optional, string
- `camlCondition` - SharePoint CAML conditions to use, optional, string
- `metaFields` - array of internal field names to request along with the files, optional, array of strings

#### Overloads / cases (checked are implemented, unchecked will be soon):
- [x] All files with folder structure from spRootFolder
- [x] Files from spRootFolder folder, first hierarchy level only
- [x] Folders structure from spRootFolder without files
- [ ] Files based on array of paths provided strictly
- [ ] Files based on REST filters conditions
- [ ] Files based on CAML query conditions
- [ ] Pull for documents metadata to use it in callback's custom logic

Use case scenarios could be found on the [Scenarios](https://github.com/koltyakov/sppull/tree/master/docs/Scenarios.md) page, it presents different variations of options usage for different situation.

#### successHandler
Callback gets called upon successful files download.

#### errorHandler
Callback gets executed in case of exception inside `sppull`. Accepts error object as first argument for callback.

## Samples
Use [Scenarios](https://github.com/koltyakov/sppull/tree/master/docs/Scenarios.md) page to see different options available with `sppull`.

#### Basic usage:
```javascript
var sppull = require("sppull").sppull;

var context = {
    siteUrl: "http://contoso.sharepoint.com/subsite",
    username: "user@contoso.com",
    password: "_Password_"  
};

var options = {
    spRootFolder: "/subsite/Shared%20Documents/Contracts",
    dlRootFolder: "./Downloads/Contracts"
};

/* 
 * All files will be downloaded from http://contoso.sharepoint.com/subsite/Shared%20Documents/Contracts folder 
 * to __dirname + /Downloads/Contracts folder.
 * Folders structure will remain original as it is in SharePoint's target folder.
*/
sppull(context, options)
    .then(function(downloadResults) {
        console.log("Files are downloaded");
        console.log("For more, please check the results", JSON.stringify(downloadResults));
    })
    .catch(function(err) {
        console.log("Core error has happened", err);
    });
```

## Inspiration and references

This project creation was inspired by [spsave](https://github.com/s-KaiNet/spsave) by [Sergei Sergeev](https://github.com/s-KaiNet) and [gulp-spsync](https://github.com/wictorwilen/gulp-spsync) by [Wictor Wilén](https://github.com/wictorwilen) projects.

SPPull depends heavily on [sp-request](https://github.com/s-KaiNet/sp-request) module and use it to send REST queries to SharePoint.