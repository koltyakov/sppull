import { join } from 'path';

import { AuthConfig as SPAuthConfigurator, IAuthConfigSettings } from 'node-sp-auth-config';
import SPPull, { ISPPullOptions, ISPPullContext } from '../../src';

const authSettings: IAuthConfigSettings = {
  configPath: './config/integration/private.spo.json'
};

new SPAuthConfigurator(authSettings).getContext()
  .then(({ siteUrl, authOptions }) => {
    const webRelativeUrl = `/${siteUrl.split('/').slice(3).join('/')}`;
    const pullContext: ISPPullContext = { siteUrl, creds: authOptions };
    const pullOptions: ISPPullOptions = {
      spBaseFolder: '/',
      spRootFolder: 'Shared%20Documents/test',
      spDocLibUrl: `${webRelativeUrl}/Shared Documents`,
      dlRootFolder: join(__dirname, 'Downloads'),
      metaFields: [ 'Title', 'CheckoutUserId' ],
      camlCondition: `
        <Geq>
          <FieldRef Name='Modified' />
          <Value Type="DateTime">
            <Today OffsetDays="-30" />
          </Value>
        </Geq>
      `
    };
    return SPPull.download(pullContext, pullOptions);
  })
  .then(console.log)
  .then((_) => console.log('Done'))
  .catch(console.log);
