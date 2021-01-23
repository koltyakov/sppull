import { AuthConfig as SPAuthConfigurator } from 'node-sp-auth-config';
import SPPull, { ISPPullOptions, ISPPullContext } from 'sppull';

new SPAuthConfigurator().getContext().then((context) => {

  const pullContext: ISPPullContext = {
    siteUrl: context.siteUrl,
    ...context.authOptions
  } as any;

  const pullOptions: ISPPullOptions = {
    spRootFolder: 'Shared%20Documents',
    dlRootFolder: './Downloads/Documents'
  };

  SPPull.download(pullContext, pullOptions);

// tslint:disable-next-line: no-console
}).catch(console.log);
