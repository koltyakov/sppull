var fs = require('fs');
var prompt = require("prompt");

var spf = spf || {};
spf.utils = spf.utils || {};

spf.utils.deleteFolderRecursive = function(path) {
    var _self = this;
    _self.deleteFolderRecursive = function(path) {
        if (fs.existsSync(path)) {
            fs.readdirSync(path).forEach(function(file, index) {
                var curPath = path + "/" + file;
                if (fs.lstatSync(curPath).isDirectory()) {
                    _self.deleteFolderRecursive(curPath);
                } else {
                    fs.unlinkSync(curPath);
                }
            });
            fs.rmdirSync(path);
        }
    };
    _self.deleteFolderRecursive(path);
};

module.exports = spf.utils;