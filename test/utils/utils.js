var fs = require('fs');
var prompt = require("prompt");
var Cpass = require("cpass");
var cpass = new Cpass();

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

spf.utils.initConfig = function(envType, configPath, callback) {
    var _self = this;
    _self.ctx = {};
    fs.exists(configPath, function(exists) {
        if (exists) {
            _self.ctx = require(__dirname + "/../../" + configPath);
        }
        _self.ctx[envType] = _self.ctx[envType] || {};

        if (_self.ctx[envType].siteUrl && _self.ctx[envType].username && _self.ctx[envType].password) {
            if (callback && typeof callback === "function") {
                _self.ctx[envType].password = cpass.decode(_self.ctx[envType].password);
                callback(_self.ctx[envType]);
            }
        } else {
            var promptFor = [];
            promptFor.push({
                description: "SharePoint Site Url",
                name: "siteUrl",
                type: "string",
                default: _self.ctx[envType].siteUrl || "",
                required: true
            });
            promptFor.push({
                description: "User login",
                name: "username",
                type: "string",
                default: _self.ctx[envType].username || "",
                required: true
            });
            if (envType === 'onprem') {
                promptFor.push({
                    description: "Domain (for On-Prem only)",
                    name: "domain",
                    type: "string",
                    default: _self.ctx[envType].domain || "",
                    required: false
                });
            }
            promptFor.push({
                description: "Password",
                name: "password",
                type: "string",
                hidden: true,
                replace: "*",
                required: true
            });
            promptFor.push({
                description: "Do you want to save config to disk?",
                name: "save",
                type: "boolean",
                default: true,
                required: true
            });
            prompt.start();
            prompt.get(promptFor, function (err, res) {
                var json = {};
                json.siteUrl = res.siteUrl;
                json.username = res.username;
                json.password = cpass.encode(res.password);
                if ((res.domain || "").length > 0) {
                    json.domain = res.domain;
                }
                _self.ctx[envType] = json;
                if (res.save) {
                    console.log(JSON.stringify(_self.ctx));
                    fs.writeFile(configPath, JSON.stringify(_self.ctx), "utf8", function(err) {
                        if (err) {
                            console.log(err);
                            return;
                        }
                        console.log("Config file is saved to " + configPath);
                    });
                }
                if (callback && typeof callback === "function") {
                    _self.ctx[envType].password = cpass.decode(_self.ctx[envType].password);
                    callback(_self.ctx[envType]);
                }
            });
        }
    });
};

module.exports = spf.utils;