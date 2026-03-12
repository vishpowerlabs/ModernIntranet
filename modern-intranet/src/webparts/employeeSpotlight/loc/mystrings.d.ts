declare interface IEmployeeSpotlightWebPartStrings {
    PropertyPaneDescription: string;
    BasicGroupName: string;
    DataSourceGroupName: string;
    ColumnMappingGroupName: string;
    DisplayGroupName: string;
    ShowTitleFieldLabel: string;
    TitleFieldLabel: string;
    ShowBackgroundBarFieldLabel: string;
    TitleBarStyleFieldLabel: string;
    TitleBarStyleSolidOption: string;
    TitleBarStyleUnderlineOption: string;
}

declare module 'EmployeeSpotlightWebPartStrings' {
    const strings: IEmployeeSpotlightWebPartStrings;
    export = strings;
}
