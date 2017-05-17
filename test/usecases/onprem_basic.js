// node ./test/usecases/onprem_basic.js

let cpass = new (require('cpass'))();
let colors = require('colors');

let sppull = require('./../../dist').sppull;
let utils = require('./../utils/utils');

let configPath = './../config/private.json';
let envType = 'onprem';

let config = require(configPath)[envType];

let context = {
    siteUrl: config.siteUrl,
    creds: {
        username: config.username,
        password: cpass.decode(config.password)
    }
};

let options = {
    spRootFolder: '_catalogs/masterpage',
    dlRootFolder: './downloads/onprem'
};

// utils.deleteFolderRecursive(options.dlRootFolder);

console.log(colors.yellow('\n=== On-Prem - Basic ===\n'));

sppull(context, options)
    .then((data) => {
        console.log(colors.green('\n=== Finished ===\n'));
    })
    .catch((err) => {
        console.log(colors.red(err));
    });
