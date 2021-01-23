import { AuthConfig, IAuthConfigSettings, IAuthContext } from 'node-sp-auth-config';
import { getAuth as getNodeAuth, IAuthResponse } from 'node-sp-auth';
import { ICiEnvironmentConfig, IPrivateEnvironmentConfig, IEnvironmentConfig } from '../configs';

export interface IConfig {
  configPath?: string;
  authConfigSettings?: {
    headlessMode: boolean;
    authOptions: IAuthContext;
  };
}

export const getAuthConf = (config: IEnvironmentConfig): IConfig => {
  const proxySettings: IConfig =
    typeof (config as IPrivateEnvironmentConfig).configPath !== 'undefined'
      ? { // Local test mode
        configPath: (config as IPrivateEnvironmentConfig).configPath
      }
      : { // Headless/CI mode
        authConfigSettings: {
          headlessMode: true,
          authOptions: {
            siteUrl: (config as ICiEnvironmentConfig).siteUrl,
            ...(config as ICiEnvironmentConfig).authOptions
          } as unknown as IAuthContext
        }
      };
  return proxySettings;
};

export const getAuth = async (config: IEnvironmentConfig): Promise<IAuthResponse> => {
  const authConf = getAuthConf(config);
  const { siteUrl, authOptions } = await new AuthConfig(({
    configPath: authConf.configPath,
    ...authConf.authConfigSettings || {}
  } as unknown as IAuthConfigSettings)).getContext();
  return await getNodeAuth(siteUrl, authOptions);
};

export const getAuthCtx = (config: IEnvironmentConfig): Promise<IAuthContext> => {
  const authConf = getAuthConf(config);
  return new AuthConfig({
    configPath: authConf.configPath,
    ...authConf.authConfigSettings || {},
    headless: true
  } as unknown as IAuthConfigSettings).getContext();
};
