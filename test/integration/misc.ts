import { AuthConfig, IAuthContext } from 'node-sp-auth-config';
import { getAuth as getNodeAuth } from 'node-sp-auth';
import { ICiEnvironmentConfig, IPrivateEnvironmentConfig, IEnvironmentConfig } from '../configs';

export const getAuthConf = (config: IEnvironmentConfig) => {
  const proxySettings =
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
        }
      }
    };
  return proxySettings;
};

export const getAuth = (config: IEnvironmentConfig) => {
  const authConf = getAuthConf(config);
  return new AuthConfig({
    configPath: authConf.configPath,
    ...authConf.authConfigSettings || {}
  }).getContext()
    .then(({ siteUrl, authOptions }) => {
      return getNodeAuth(siteUrl, authOptions);
    });
};

export const getAuthCtx = (config: IEnvironmentConfig): Promise<IAuthContext> => {
  const authConf = getAuthConf(config);
  return new AuthConfig({
    configPath: authConf.configPath,
    ...authConf.authConfigSettings || {}
  }).getContext();
};
