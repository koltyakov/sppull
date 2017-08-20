export interface ITestSetup {
    environmentName: string;
    configPath: string;
    spRootFolder: string;
    dlRootFolder: string;
}

export const TestsConfigs: ITestSetup[] = [
    {
        environmentName: 'SharePoint Online',
        configPath: './config/integration/private.spo.json',
        spRootFolder: 'Shared Documents',
        dlRootFolder: './downloads/spo'
    }, {
        environmentName: 'On-Premise 2016',
        configPath: './config/integration/private.2016.json',
        spRootFolder: 'Shared Documents',
        dlRootFolder: './downloads/2016'
    }
    // , {
    //     environmentName: 'On-Premise 2013',
    //     configPath: './config/integration/private.2013.json',
    //     spRootFolder: 'Shared Documents',
    //     dlRootFolder: './downloads/2013'
    // }
];
