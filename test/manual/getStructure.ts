import { AuthConfig as SPAuthConfigurator, IAuthConfigSettings } from 'node-sp-auth-config';

import { ISPPullOptions, ISPPullContext } from '../../src';
import PullAPI from '../../src/api';
import { IContent } from '../../src/interfaces/content';

const authSettings: IAuthConfigSettings = {
  configPath: './config/integration/private.spo.json'
};

const getStructure = async (pull: PullAPI, folder: string, data: IContent = { folders: [], files: [] }): Promise<IContent> => {
  const folderContent = await pull.getFolderContent(folder);
  data.folders = data.folders.concat(folderContent.folders);
  data.files = data.files.concat(folderContent.files);
  for (const folder of folderContent.folders) {
    console.log(`== Getting data: ${folder.ServerRelativeUrl} ==`);
    await getStructure(pull, folder.ServerRelativeUrl, data);
  }
  return data;
};

new SPAuthConfigurator(authSettings).getContext()
  .then(({ siteUrl, authOptions }) => {
    const ctx: ISPPullContext = { siteUrl, ...authOptions } as any;
    const opts: ISPPullOptions = {
      spBaseFolder: '/',
      spRootFolder: 'Shared Documents'
    };
    const pull = new PullAPI(ctx, opts);
    return getStructure(pull, opts.spRootFolder);
  })
  .then((content) => {
    const data = {
      folders: content.folders.map(({ ServerRelativeUrl, ItemCount }) => ({ ServerRelativeUrl, ItemCount })),
      files: content.files.map(({ ServerRelativeUrl, TimeLastModified, Length }) => ({ ServerRelativeUrl, TimeLastModified, Length }))
    };
    console.log(JSON.stringify(data, null, 2));
  })
  .catch(console.log);
