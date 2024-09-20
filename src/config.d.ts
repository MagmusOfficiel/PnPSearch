declare interface IGlobalSetting {
    Domain: string,
    BaseUrl: string,
    SearchBoxId: string,
    SearchBoxInstanceId: string,
    SearchResultId: string,
    SearchResultInstanceId: string,
    SearchFilterId: string,
    SearchFilterInstanceId: string,
    SearchVerticalId: string,
    SearchVerticalInstanceId: string,
    FolderConfig: string,
    ApiWeb: string,
}

declare module 'GlobalSetting' {
const configs: IGlobalSetting;
export = configs;
}
