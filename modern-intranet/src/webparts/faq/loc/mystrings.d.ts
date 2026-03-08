declare interface IFaqWebPartStrings {
    PropertyPaneDescription: string;
    BasicGroupName: string;
    DataSourceGroupName: string;
    ColumnMappingsGroupName: string;
    DisplaySettingsGroupName: string;

    ShowTitleFieldLabel: string;
    TitleFieldLabel: string;
    ShowBackgroundBarFieldLabel: string;
    TitleBarStyleFieldLabel: string;
    TitleBarStyleSolidOption: string;
    TitleBarStyleUnderlineOption: string;

    SiteUrlFieldLabel: string;
    ListIdFieldLabel: string;

    QuestionColumnFieldLabel: string;
    AnswerColumnFieldLabel: string;
    CategoryColumnFieldLabel: string;
    OrderColumnFieldLabel: string;

    ShowSearchFieldLabel: string;
    ShowCategoryFilterFieldLabel: string;
    AllowMultipleOpenFieldLabel: string;
    ExpandFirstItemFieldLabel: string;
}

declare module 'FaqWebPartStrings' {
    const strings: IFaqWebPartStrings;
    export = strings;
}
