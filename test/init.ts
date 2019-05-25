import { resolve } from 'path';
import { AuthConfig } from 'node-sp-auth-config';
import * as colors from 'colors';

import { Environments, IPrivateEnvironmentConfig } from './configs';

async function checkOrPromptForIntegrationConfigCreds (): Promise<void> {

  for (const conf of Environments) {
    console.info(`\n=== ${colors.bold.yellow(`${conf.environmentName} Credentials`)} ===\n`);
    const c = conf as IPrivateEnvironmentConfig;
    if (typeof c.configPath !== 'undefined') {
      await new AuthConfig({ configPath: c.configPath }).getContext();
      console.info(colors.grey(`Gotcha ${resolve(c.configPath)}`));
    } else {
      console.info(colors.yellow(`CI configuration detected.`));
    }
  }

  console.info('\n');

}

checkOrPromptForIntegrationConfigCreds();
