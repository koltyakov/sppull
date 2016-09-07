// node ./test/usecases/online_strict.js

var sppull = require("./../../lib/src/index").sppull;
var utils = require("./../utils/utils");
var colors = require("colors");

var configPath = "test/config/config.json";
var envType = "online";

utils.initConfig(envType, configPath, function(config) {

    var context = {
        siteUrl: config.siteUrl,
        username: config.username,
        password: config.password
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

    // utils.deleteFolderRecursive(options.dlRootFolder);

    console.log(colors.yellow("\n=== Online - Strinct objects ===\n"));

    sppull(context, options)
        .then(function(data) {
            console.log(colors.green("\n=== Finished ===\n"));
        })
        .catch(function(err) {
            console.log(colors.red(err));
        });

});