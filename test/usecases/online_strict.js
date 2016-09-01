// node ./test/usecases/online_strict.js

var sppull = require("./../../lib/src/index").sppull;
var config = require("./../config/config");
var utils = require("./../utils/utils");
var colors = require("colors");

var context = {
    siteUrl: config.online.siteUrl,
    username: config.online.username,
    password: config.online.password
};

var options = {
    spRootFolder: "_catalogs/masterpage",
    dlRootFolder: "./downloads/online",
    strictObjects: [
        "seattle.master",
        "/oslo.master",
        "v4.master"
    ]
};

utils.deleteFolderRecursive(options.dlRootFolder);

console.log(colors.yellow("\n=== Online - Basic ===\n"));

sppull(context, options)
    .then(function(data) {
        console.log(colors.green("\n=== Finished ===\n"));
    })
    .catch(function(err) {
        console.log(colors.red(err));
    });