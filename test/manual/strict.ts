import { join } from 'path';

import { AuthConfig as SPAuthConfigurator, IAuthConfigSettings } from 'node-sp-auth-config';
import { ISPPullOptions, ISPPullContext, Download as IDownload } from './../../src';

const Download: IDownload = require(join(__dirname, '../../src'));
const sppull = Download.sppull;

const authSettings: IAuthConfigSettings = {
  configPath: './config/integration/private.spo.json'
};

new SPAuthConfigurator(authSettings).getContext()
  .then(({ siteUrl, authOptions }) => {
    const pullContext: ISPPullContext = { siteUrl, ...authOptions } as any;
    const pullOptions: ISPPullOptions = {
      spRootFolder: "",
      dlRootFolder: join(__dirname, 'Downloads'),
      strictObjects: [
        `/sites/ci/Shared Documents/gosip's.png`
      ]
    };
    return sppull(pullContext, pullOptions);
  })
  .then(_ => console.log('Done'))
  .catch(console.log);
