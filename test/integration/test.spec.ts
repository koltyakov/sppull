import { expect } from 'chai';
import * as path from 'path';
import { Cpass } from 'cpass';

import { ISPPullOptions, ISPPullContext, Download } from '../../src/SPPull';
const sppull = (new Download).sppull;
const cpass = new Cpass();

import { deleteFolderRecursive } from '../utils/utils';
import { TestsConfigs } from '../configs';

for (let testConfig of TestsConfigs) {

    describe(`Run tests in ${testConfig.environmentName}`, () => {

        let context: ISPPullContext;
        let config: any;

        before('Configure', function(done: any): void {
            this.timeout(30 * 1000);
            config = require(path.resolve(testConfig.configPath));
            context = {
                siteUrl: config.siteUrl,
                creds: {
                    ...config,
                    password: config.password && cpass.decode(config.password)
                }
            };
            done();
        });

        it(`should pull in basic mode`, function(done: MochaDone): void {
            this.timeout(300 * 1000);

            let options: ISPPullOptions = {
                spRootFolder: testConfig.spRootFolder,
                dlRootFolder: path.join(testConfig.dlRootFolder, 'basic')
            };

            sppull(context, options)
                .then((data) => {
                    done();
                })
                .catch(done);
        });

        it(`should pull in strict mode`, function(done: MochaDone): void {
            this.timeout(300 * 1000);

            let options: ISPPullOptions = {
                spRootFolder: '_catalogs/masterpage',
                dlRootFolder: path.join(testConfig.dlRootFolder, 'strict'),
                strictObjects: [
                    'seattle.master',
                    '/oslo.master',
                    'v4.master'
                ]
            };

            sppull(context, options)
                .then((data) => {
                    done();
                })
                .catch(done);
        });

        after('Deleting test objects', function(done: MochaDone): void {
            this.timeout(150 * 1000);
            // deleteFolderRecursive(testConfig.dlRootFolder);
            done();
        });

    });

}
