# Usage scenarios

## SPPull initiation
sppull module object should be initiated by requiring node.js module or sources reference first.
```javascript
var sppull = require("sppull").sppull;
// sppull then could be used as a downloading client
```

## Context and authentication
The context object is used to define target SharePoint site and user's credentials.
It's recommended to take `username` and `password` out of Git stored code.
Credentials could be stored in process environment variables or any config which is listed in .gitignore exceptions, for instance.   
```javascript
var context = {
    siteUrl: "http://contoso.sharepoint.com/subsite",
    username: "user@contoso.com",
    password: "_Password_"  
};
```

## Basic usage
### Downloaded all files keeping target source folder structure
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

sppull(context, options)
    .then(function(downloadResults) {
        console.log("Files are downloaded");
        console.log("For more, please check the results", JSON.stringify(downloadResults));
    })
    .catch(function(err) {
        console.log("Core error has happened", err);
    });
```
All examples use the same initiation and it's omitted from the samples.
Actually if you know options parameters you know all sppull, at least by now.

### Downloaded files from target source folder explicitly
```javascript
var options = {
    spRootFolder: "/subsite/Shared%20Documents/Contracts",
    dlRootFolder: "./Downloads/Contracts",
    recursive: false
};
```

### Recreate locally folders' structure from SharePoint target
```javascript
var options = {
    spRootFolder: "/subsite/Shared%20Documents/Contracts",
    dlRootFolder: "./Downloads/Contracts",
    foderStructureOnly: false
};
```

### Download specific files or folders within root folder explicitly

> Works only with array of files names, folders aren't implemented yet

```javascript
var options = {
    spRootFolder: "/subsite/_catalogs/masterpage",
    dlRootFolder: "./Downloads/Assets",
    strictObjects: [
        "/Display%20Templates/Search", // All files and folders under /subsite/_catalogs/masterpage/Display Templates/Search folder
        "/Display%20Templates/Filters/Filter_Slider.js", // Only Filter_Slider.js from /subsite/_catalogs/masterpage/Display Templates/Filters
        "/responsive.master" // Only responsive.master file from /subsite/_catalogs/masterpage
    ]
};
```
If a file is sitting within explicitly provided folder, folder settings will be used, as a file setting is a subset of it's parent.

### Download files from spRootFolder ignoring some folders and files

> Not implemented yet

```javascript
var options = {
    spRootFolder: "/subsite/_catalogs/masterpage",
    dlRootFolder: "./Downloads/Assets",
    excludeObjects: [
        "/Display%20Templates/Search",
        "/Display%20Templates/Filters/somefiletoignore.js",
        "/somelargefile.mp4"
    ]
};
```

### Download files which correspond to a REST filter condition

> Not implemented yet

```javascript
// REST filter for Approved documents
var restFilters = "$filter=OData__ModerationStatus eq 0";
var options = {
    spRootFolder: "/subsite/Shared%20Documents/Contracts",
    dlRootFolder: "./Downloads/Contracts",
    restCondition: restFilters
};
```

### Download files which correspond to a CAML condition

> Not implemented yet

```javascript
// CAML to filter only Approved documents
var camlString = "<Query>" +
                    "<Where>" +
                        "<Eq>" +
                        "<FieldRef Name='_ModerationStatus' />" +
                        "<Value Type='ModStat'>0</Value>" +
                        "</Eq>" +
                    "</Where>" +
                 "</Query>";
var options = {
    spRootFolder: "/subsite/Shared%20Documents/Contracts",
    dlRootFolder: "./Downloads/Contracts",
    camlCondition: camlString
};
```

### Download files with their metadata

> Not implemented yet, but basic metadata is already accessible in the callback <br>
> not in results[...].metadata though

```javascript
var options = {
    spRootFolder: "/subsite/Shared%20Documents/Contracts",
    dlRootFolder: "./Downloads/Contracts",
    metaFields: [
        "Title",
        "_ModerationStatus",
        "Modified",
        "Editor"
    ]
};

sppull(context, options)
    .then(function(downloadResults) {
        console.log("Files are downloaded");
        downloadResults.forEach(function(file) {
            console.log("File " + file.ServerRelativeUrl +  " is downloaded to " + file.SavedToPath);
            console.log("It's pulled metadata is", file.metadata);
        });
    })
    .catch(function(err) {
        console.log("Core error has happened", err);
    });
```