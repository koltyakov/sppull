# SPPull - simple client to pull and download files from SharePoint

[![NPM](https://nodei.co/npm/sppull.png?mini=true&downloads=true&downloadRank=true&stars=true)](https://nodei.co/npm/sppull/)

[![npm version](https://badge.fury.io/js/sppull.svg)](https://badge.fury.io/js/sppull)
[![Downloads](https://img.shields.io/npm/dm/sppull.svg)](https://www.npmjs.com/package/sppull)
![Build Status](https://koltyakov.visualstudio.com/SPNode/_apis/build/status/sppull?branchName=master)
[![Gitter](https://badges.gitter.im/gitterHQ/gitter.png)](https://gitter.im/sharepoint-node/Lobby)

Node.js module for downloading files from SharePoint document libraries.

## New in version 2.7.0

- [Breaking change]: Removed legacy module.exports for straightforward usage in modern JS and TypeScript

```typescript
import SPPull from 'sppull';
const sppull = SPPull.download; // <-- this is an entry point
```

- TSLint replaced with ESLint
- Integration tests are migrated to new version of Mocha and fixed

## New in version 2.2.0

Files streaming download:

- Download files of any supported size
- Effective memmory consumption while fetching large files

## New in version 2.1.0

Performance in SPO and HTTPS environments is improved. Download on multiple objects is x2 faster now!

## New in version 2.0.6

Smart re-download mechanism. Existing files with no changes are ignored from the download.

## Supported SharePoint versions

- SharePoint Online
- SharePoint 2013
- SharePoint 2016
- SharePoint 2019

## How to use

### Install

```bash
npm install sppull --save-dev
```

### Demo

![How it works](http://koltyakov.ru/images/sppull-demo.gif)

### Usage

```javascript
const { SPPull } = require('sppull');

const context = {/*...*/};
const options = {/*...*/};

SPPull.download(context, options)
  .then(successHandler)
  .catch(errorHandler);
```

#### Arguments

##### Context

- `siteUrl` - SharePoint site (SPWeb) url [string, required]
- `creds`
  - `username` - user name for SP authentication [string, optional in case of some auth methods]
  - `password` - password [string, optional in case of some auth methods]
  ...

**Additional authentication options:**

Since communication module (sp-request), which is used in sppull, had received additional SharePoint authentication methods, they are also supported in sppull.

- SharePoint On-Premise (Add-In permissions)
  - `clientId`
  - `issuerId`
  - `realm`
  - `rsaPrivateKeyPath`
  - `shaThumbprint`
- SharePoint On-Premise (NTLM handshake - more commonly used scenario):
  - `username` - username without domain
  - `password`
  - `domain` / `workstation`
- SharePoint Online (Add-In permissions):
  - `clientId`
  - `clientSecret`
- SharePoint Online (SAML based with credentials - more commonly used scenario):
  - `username` - user name for SP authentication [string, required]
  - `password` - password [string, required]
- ADFS user credantials:
  - `username`
  - `password`
  - `relyingParty`
  - `adfsUrl`

For more information please check node-sp-auth [credential options](https://github.com/s-KaiNet/node-sp-auth#params) and [wiki pages](https://github.com/s-KaiNet/node-sp-auth/wiki).

##### Options

- `spRootFolder` - root folder in SharePoint to pull from [string, required]
- `dlRootFolder` - local root folder where files and folders will be saved to [string, optional, default: `./Downloads`]
- `spBaseFolder` - base folder path which is omitted then saving files locally [string, optional, default: equals to spRootFolder]
- `recursive` - to pull all files and folders recursively [boolean, optional, default: `true`]
- `ignoreEmptyFolders` - to ignore local creation of SharePoint empty folders [boolean, optional, default: `true`]
- `foderStructureOnly` - to ignore files, recreate only folders' structure [boolean, optional, default: `false`]
- `strictObjects` - array of files and folders relative paths within the `spRootFolder` to proceed explicitly, [array of strings, optional]
- `camlCondition` - SharePoint CAML conditions to use [string, optional]
- `spDocLibUrl` - SharePoint document library URL [string, mandatory with `camlCondition`]
- `metaFields` - array of internal field names to request along with the files [array of strings, optional]
- `createEmptyFolders` - to create empty folders along with documents download task [boolean, optional, default: `true`]
- `omitFolderPath` - folder path pattern which is omitted from final downloaded files path [string, optional]
- `muteConsole` - to mute console messages during transport queries to SharePoint API [boolean, optional, default: `false`]
- `shouldDownloadFile` - a callback to decide if a file should be downloaded based on the metadata [optional, (props) => boolean, default: undefined]

#### Overloads / cases

- All files with folder structure from spRootFolder
- Files from spRootFolder folder, first hierarchy level only
- Folders structure from spRootFolder without files
- Files based on array of paths provided strictly [works with array of files only right now]
- Files based on CAML query conditions
- Pull for documents metadata to use it in callback's custom logic

Use case scenarios can be found on the [Scenarios](https://github.com/koltyakov/sppull/tree/master/docs/Scenarios.md) page. This page suggests combinations of options which are optimal for certain use cases.

#### successHandler

Callback gets called upon successful files download.

#### errorHandler

Callback gets executed in case of exception inside `sppull`. Accepts error object as first argument for callback.

## Samples

Refer to the [Scenarios](https://github.com/koltyakov/sppull/tree/master/docs/Scenarios.md) page for suggested options combinations available with `sppull`.

### Basic usage

TypeSctipt:

```typescript
import { AuthConfig as SPAuthConfigurator } from 'node-sp-auth-config';
import SPPull, { ISPPullOptions, ISPPullContext } from 'sppull';

new SPAuthConfigurator().getContext().then(({ siteUrl, authOptions }) => {

  const context: ISPPullContext = {
    siteUrl,
    ...authOptions
  } as any;

  const options: ISPPullOptions = {
    spRootFolder: 'Shared%20Documents',
    dlRootFolder: './Downloads/Documents'
  };

  SPPull.download(context, options);

}).catch(console.log);
```

or ES6:

```javascript
const { SPPull } = require('sppull');

const context = {
  siteUrl: "http://contoso.sharepoint.com/subsite",
  creds: {
    username: "user@contoso.com",
    password: "_Password_",
    online: true
  }
};

const options = {
  spRootFolder: "Shared%20Documents/Contracts",
  dlRootFolder: "./Downloads/Contracts"
};

/*
 * All files will be downloaded from http://contoso.sharepoint.com/subsite/Shared%20Documents/Contracts folder
 * to __dirname + /Downloads/Contracts folder.
 * Folders structure will remain original as it is in SharePoint's target folder.
*/
SPPull.download(context, options)
  .then((downloadResults) => {
    console.log("Files are downloaded");
    console.log("For more, please check the results", JSON.stringify(downloadResults));
  })
  .catch((err) => {
    console.log("Core error has happened", err);
  });
```

### Passwords storage

To eliminate any local password storing if preferable to use any two-way hashing technique, like [cpass](https://github.com/koltyakov/cpass).

## Inspiration and references

This project was inspired by [spsave](https://github.com/s-KaiNet/spsave) by [Sergei Sergeev](https://github.com/s-KaiNet) and [gulp-spsync](https://github.com/wictorwilen/gulp-spsync) by [Wictor Wilén](https://github.com/wictorwilen) projects.

SPPull depends heavily on [sp-request](https://github.com/s-KaiNet/sp-request) module and use it to send REST queries to SharePoint.
