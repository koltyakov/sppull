// import { expect } from 'chai';
import * as path from 'path';
import * as fs from 'fs';
import { Cpass } from 'cpass';

import SPPull, { ISPPullOptions, ISPPullContext } from '../../src';

const cpass = new Cpass();

import { deleteFolderRecursive } from '../utils/utils';
import { Environments } from '../configs';
import { getAuthCtx } from './misc';

for (const testConfig of Environments) {

  describe(`Run tests in ${testConfig.environmentName}`, () => {

    let context: ISPPullContext;

    before('Configure', function(done: Mocha.Done): void {
      this.timeout(30 * 1000);

      getAuthCtx(testConfig)
        .then((ctx) => {
          const { siteUrl, authOptions } = ctx;
          interface WPass { password?: string; }
          if ((authOptions as WPass).password) {
            (authOptions as WPass).password = cpass.decode((authOptions as WPass).password);
          }
          context = { siteUrl, creds: authOptions };
          deleteFolderRecursive(testConfig.dlRootFolder);
          done();
        })
        .catch(done);
    });

    it('should pull in basic mode', function(done: Mocha.Done): void {
      this.timeout(300 * 1000);
      SPPull.download(context, {
        spRootFolder: testConfig.spRootFolder,
        dlRootFolder: path.join(testConfig.dlRootFolder, 'basic'),
        muteConsole: true
      }).then(() => done()).catch(done);
    });

    it('should pull in strict mode', function(done: Mocha.Done): void {
      this.timeout(300 * 1000);
      SPPull.download(context, {
        spRootFolder: '_catalogs/masterpage',
        dlRootFolder: path.join(testConfig.dlRootFolder, 'strict'),
        strictObjects: [
          'seattle.master',
          '/oslo.master',
          'v4.master'
        ],
        muteConsole: true
      }).then(() => done()).catch(done);
    });

    it('should pull without subfolders data', function(done: Mocha.Done): void {
      this.timeout(100 * 1000);
      SPPull.download(context, {
        spRootFolder: '_catalogs/masterpage',
        dlRootFolder: path.join(testConfig.dlRootFolder, 'flat'),
        recursive: false,
        createEmptyFolders: false,
        muteConsole: true
      }).then(() => done()).catch(done);
    });

    it('should pull folders structure', function(done: Mocha.Done): void {
      this.timeout(300 * 1000);
      SPPull.download(context, {
        spRootFolder: '_catalogs/masterpage',
        dlRootFolder: path.join(testConfig.dlRootFolder, 'structure'),
        foderStructureOnly: true,
        muteConsole: true
      }).then(() => done()).catch(done);
    });

    it('should pull using caml condition', function(done: Mocha.Done): void {
      this.timeout(100 * 1000);

      const d = new Date();
      d.setDate(d.getDate() - 5);
      const camlString = `
        <Eq>
          <FieldRef Name='Modified' />
          <Value Type='DateTime'>${d.toISOString()}</Value>
        </Eq>
      `;

      SPPull.download(context, {
        dlRootFolder: path.join(testConfig.dlRootFolder, 'caml'),
        spDocLibUrl: 'Shared Documents',
        camlCondition: camlString,
        muteConsole: true
      }).then(() => done()).catch(done);
    });

    it('should pull with additional metadata', function(done: Mocha.Done): void {
      this.timeout(300 * 1000);

      const options: ISPPullOptions = {
        spRootFolder: '_catalogs/masterpage',
        dlRootFolder: path.join(testConfig.dlRootFolder, 'metadata'),
        recursive: false,
        metaFields: [ 'Title', 'Modified', 'Editor' ],
        muteConsole: true
      };

      SPPull.download(context, options)
        .then((data) => {
          interface WMeta { metadata?: unknown; }
          fs.writeFileSync(
            path.join(options.dlRootFolder, 'metadata.json'),
            JSON.stringify(data.map((d) => (d as WMeta).metadata), null, 2)
          );
          done();
        })
        .catch(done);
    });

    after('Deleting test objects', function(done: Mocha.Done): void {
      this.timeout(150 * 1000);
      deleteFolderRecursive(testConfig.dlRootFolder);
      done();
    });

  });

}
