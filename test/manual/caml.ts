import { join } from 'path';

import { AuthConfig as SPAuthConfigurator, IAuthConfigSettings } from 'node-sp-auth-config';
import { ISPPullOptions, ISPPullContext, Download as IDownload } from '../../src';

const Download: IDownload = require(join(__dirname, '../../src'));
const sppull = Download.sppull;

const authSettings: IAuthConfigSettings = {
  configPath: './config/integration/private.spo.json'
};

new SPAuthConfigurator(authSettings).getContext()
  .then(({ siteUrl, authOptions }) => {
    const webRelativeUrl = `/${siteUrl.split('/').slice(3).join('/')}`;
    const pullContext: ISPPullContext = { siteUrl, ...authOptions } as any;
    const pullOptions: ISPPullOptions = {
      spBaseFolder: '/',
      // spRootFolder: 'Shared%20Documents',
      spDocLibUrl: `${webRelativeUrl}/Shared Documents`,
      dlRootFolder: join(__dirname, 'Downloads'),
      metaFields: [ 'Title' ],
      camlCondition: `
        <Geq>
          <FieldRef Name='Modified' />
          <Value Type="DateTime">
            <Today OffsetDays="-30" />
          </Value>
        </Geq>
      `
    };
    return sppull(pullContext, pullOptions);
  })
  .then(console.log)
  .then(_ => console.log('Done'))
  .catch(console.log);
