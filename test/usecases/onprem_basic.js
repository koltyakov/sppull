// node ./test/usecases/onprem_basic.js

var sppull = require("./../../lib/src/index").sppull;
var utils = require("./../utils/utils");
var colors = require("colors");

var configPath = "test/config/config.json";
var envType = "onprem";

utils.initConfig(envType, configPath, function(config) {

    var context = {
        siteUrl: config.siteUrl,
        username: config.username,
        password: config.password
    };
    if (config.domain) {
        context.domain = config.domain;
    }

    var options = {
        spRootFolder: "_catalogs/masterpage",
        dlRootFolder: "./downloads/onprem"
    };

    // utils.deleteFolderRecursive(options.dlRootFolder);

    console.log(colors.yellow("\n=== On-Prem - Basic ===\n"));

    sppull(context, options)
        .then(function(data) {
            console.log(colors.green("\n=== Finished ===\n"));
        })
        .catch(function(err) {
            console.log(colors.red(err));
        });

});