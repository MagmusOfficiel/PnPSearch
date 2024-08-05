declare interface IGlobalSetting {
    Domain: string,
    BaseUrl: string,
    SearchBoxId: string,
    SearchResultId: string,
    SearchFilterId: string,
    SearchVerticalId: string,
    FolderConfig: string,
    ApiWeb: string,
}

declare module 'GlobalSetting' {
const configs: IGlobalSetting;
export = configs;
}
