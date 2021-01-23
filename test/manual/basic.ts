import { join } from 'path';

import { AuthConfig as SPAuthConfigurator, IAuthConfigSettings } from 'node-sp-auth-config';
import SPPull, { ISPPullOptions, ISPPullContext } from './../../src';

const authSettings: IAuthConfigSettings = {
  configPath: './config/integration/private.spo.json'
};

new SPAuthConfigurator(authSettings).getContext()
  .then(({ siteUrl, authOptions }) => {
    const pullContext: ISPPullContext = { siteUrl, creds: authOptions };
    const pullOptions: ISPPullOptions = {
      spRootFolder: 'Shared Documents',
      dlRootFolder: join(__dirname, 'Downloads')
    };
    return SPPull.download(pullContext, pullOptions);
  })
  .then(() => console.log('Done'))
  .catch(console.log);
