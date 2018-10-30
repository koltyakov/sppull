import { AuthConfig } from 'node-sp-auth-config';
import * as path from 'path';
import * as colors from 'colors';

import { TestsConfigs } from './configs';

async function checkOrPromptForIntegrationConfigCreds (): Promise<void> {

  for (const { configPath, environmentName } of TestsConfigs) {
    console.log(`\n=== ${colors.bold.yellow(`${environmentName} Credentials`)} ===\n`);
    await new AuthConfig({ configPath }).getContext();
    console.log(colors.grey(`Gotcha ${path.resolve(configPath)}`));
  }

  console.log('\n');

}

checkOrPromptForIntegrationConfigCreds();
