export class Utils {

    // public checkNestedProperties(object: any, ...args: string[]): boolean {
    //     for (let i: number = 0, len: number = args.length; i < len; i += 1) {
    //         if (!object || !object.hasOwnProperty(args[i])) {
    //             return false;
    //         }
    //         object = object[args[i]];
    //     }
    //     return true;
    // }

    // public getCaseInsensitiveProp(object: any, propertyName: string): any {
    //     propertyName = propertyName.toLowerCase();
    //     return Object.keys(object).reduce((res: any, prop: string) => {
    //         if (prop.toLowerCase() === propertyName) {
    //             res = object[prop];
    //         }
    //         return res;
    //     }, undefined);
    // }

    public isOnPrem(url: string): boolean {
        return url.indexOf('.sharepoint.com') === -1 && url.indexOf('.sharepoint.cn') === -1;
    }

    public isUrlHttps(url: string): boolean {
        return url.split('://')[0].toLowerCase() === 'https';
    }

    public isUrlAbsolute(url: string): boolean {
        return url.indexOf('http:') === 0 || url.indexOf('https:') === 0;
    }

    public combineUrl(...args: string[]): string {
        return args.join('/').replace(/(\/)+/g, '/').replace(':/', '://');
    }

    // public mergeHeaders(...args: any[]): Headers {
    //     return args.reduce((headers: Headers, headersPatch: any) => {
    //         this.anyToHeaders(headersPatch).forEach((value: string, name: string) => {
    //             headers.set(name, value);
    //         });
    //         return headers;
    //     }, new Headers());
    // }

    // public anyToHeaders(headers: any = {}): Headers {
    //     return (new Request('', { headers })).headers;
    // }

}
