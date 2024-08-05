declare interface ISearchWebPartStrings {
        SelectConfiguration: string,
        ReloadConfiguration: string,
        ErrorLoadingConfigs: string,
        ErrorRendering: string
}

declare module 'SearchWebPartStrings' {
    const strings: ISearchWebPartStrings;
    export = strings;
}
