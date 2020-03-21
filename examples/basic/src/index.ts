import { AuthConfig as SPAuthConfigurator } from 'node-sp-auth-config';
import { ISPPullOptions, ISPPullContext, Download as IDownload } from 'sppull';

new SPAuthConfigurator().getContext().then((context) => {

  const Download: IDownload = require('sppull');
  const sppull = Download.sppull;

  const pullContext: ISPPullContext = {
    siteUrl: context.siteUrl,
    ...context.authOptions
  } as any;

  const pullOptions: ISPPullOptions = {
    spRootFolder: 'Shared%20Documents',
    dlRootFolder: './Downloads/Documents'
  };

  sppull(pullContext, pullOptions);

// tslint:disable-next-line: no-console
}).catch(console.log);
