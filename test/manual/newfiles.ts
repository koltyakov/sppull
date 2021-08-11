import { join } from 'path';

import { AuthConfig as SPAuthConfigurator, IAuthConfigSettings } from 'node-sp-auth-config';
import SPPull, { ISPPullOptions, ISPPullContext } from '../../src';

const authSettings: IAuthConfigSettings = {
  configPath: './config/integration/private.spo.json'
};

const downloadSince = new Date(); // persist/restore a date time in UTC of the the last sync start session
downloadSince.setDate(downloadSince.getDate() - 1); 

new SPAuthConfigurator(authSettings).getContext()
  .then(({ siteUrl, authOptions }) => {
    const pullContext: ISPPullContext = { siteUrl, creds: authOptions };
    const pullOptions: ISPPullOptions = {
      spRootFolder: 'Shared Documents',
      dlRootFolder: join(__dirname, 'Downloads'),
      muteConsole: true,
      // Use CAML for libs with less than 5000 documents
      camlCondition: `
        <Geq>
          <FieldRef Name='Modified' />
          <Value Type="DateTime">
            <Today OffsetDays="-1" />
          </Value>
        </Geq>
      `,
      // Use should download callback for large libraries
      shouldDownloadFile: (file) => {
        if (file.TimeLastModified) {
          const lastMod = new Date(file.TimeLastModified);
          return lastMod >= downloadSince;
        }
        return true;
      }
    };
    return SPPull.download(pullContext, pullOptions);
  })
  .then(() => console.log('Done'))
  .catch(console.log);
