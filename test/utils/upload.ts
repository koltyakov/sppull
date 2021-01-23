import { AuthConfig, IAuthContext } from 'node-sp-auth-config';
import { spsave } from 'spsave';
import * as fs from 'fs';
import * as path from 'path';

interface IPublishProfile {
  src: string;
  dest: string;
}

const authConfig = new AuthConfig({
  configPath: './config/private.json',
  encryptPassword: true,
  saveConfigOnDisk: true
});

const structureFolder = path.join(__dirname, '../structure');
const publishProfiles: IPublishProfile[] = fs.readdirSync(structureFolder)
  .filter((objPath) => {
    const folderPath = path.join(structureFolder, objPath);
    return fs.statSync(folderPath).isDirectory();
  })
  .map((folder) => {
    return {
      src: path.join(structureFolder, folder),
      dest: folder
    } as IPublishProfile;
  });

const walkSync = (dir: string, filelist: string[]): string[] => {
  const files = fs.readdirSync(dir);
  filelist = filelist || [];
  files.forEach((file) => {
    const filePath = path.join(dir, file);
    if (fs.statSync(filePath).isDirectory()) {
      filelist = walkSync(filePath, filelist);
    } else {
      filelist.push(filePath);
    }
  });
  return filelist;
};

const publishAll = async (profiles: IPublishProfile[], context: IAuthContext): Promise<string> => {
  const coreOptions = {
    siteUrl: context.siteUrl,
    notification: false,
    checkin: true,
    checkinType: 1
  };
  for (const profile of profiles) {
    const files = walkSync(profile.src, []);
    console.log(`=== Publishing "${profile.dest}" ===`);
    for (const file of files) {
      const fileOptions = {
        folder: `${profile.dest}/${path.dirname(path.relative(profile.src, file)).replace(/\\/g, '/')}`,
        fileName: path.basename(file),
        fileContent: fs.readFileSync(file)
      };
      await Promise.resolve(spsave(coreOptions, context.authOptions, fileOptions));
    }
  }
  return 'Done';
};

authConfig.getContext()
  .then((ctx) => publishAll(publishProfiles, ctx))
  .then(console.log)
  .catch(console.log);
