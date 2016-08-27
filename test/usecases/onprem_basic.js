// node ./test/usecases/onprem_basic.js

var sppull = require("./../../lib/src/index").sppull;
var config = require("./../config/config");
var utils = require("./../utils/utils");
var colors = require("colors");

var context = {
    siteUrl: config.onprem.siteUrl,
    username: config.onprem.username,
    password: config.onprem.password
};

if (config.onprem.domain) {
    context.domain = config.onprem.domain;
}

var options = {
    spRootFolder: "_catalogs/masterpage",
    dlRootFolder: "./downloads/onprem"
};

utils.deleteFolderRecursive(options.dlRootFolder);

console.log(colors.yellow("\n=== On-Prem - Basic ===\n"));

sppull(context, options)
    .then(function(data) {
        console.log(colors.green("\n=== Finished ===\n"));
    })
    .catch(function(err) {
        console.log(colors.red(err));
    });