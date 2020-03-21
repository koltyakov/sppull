"use strict";
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
Object.defineProperty(exports, "__esModule", { value: true });
var node_sp_auth_config_1 = require("node-sp-auth-config");
new node_sp_auth_config_1.AuthConfig().getContext().then(function (context) {
    var Download = require('sppull');
    var sppull = Download.sppull;
    var pullContext = __assign({ siteUrl: context.siteUrl }, context.authOptions);
    var pullOptions = {
        spRootFolder: 'Shared%20Documents',
        dlRootFolder: './Downloads/Documents'
    };
    sppull(pullContext, pullOptions);
}).catch(console.log);
//# sourceMappingURL=index.js.map