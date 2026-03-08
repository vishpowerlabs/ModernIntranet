/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

declare interface IHighlightsWebPartStrings {
    PropertyPaneDescription: string;
    DataSourceGroupName: string;
    ColumnMappingGroupName: string;
    DisplaySettingsGroupName: string;
    SiteUrlFieldLabel: string;
    ListIdFieldLabel: string;
    TitleColumnFieldLabel: string;
    DescriptionColumnFieldLabel: string;
    BannerImageColumnFieldLabel: string;
    LinkColumnFieldLabel: string;
    PinnedColumnFieldLabel: string;
    MaxItemsFieldLabel: string;
    ColumnsFieldLabel: string;
    ShowTitleFieldLabel: string;
    TitleFieldLabel: string;
    ShowBackgroundBarFieldLabel: string;
    TitleBarStyleFieldLabel: string;
    TitleBarStyleSolidOption: string;
    TitleBarStyleUnderlineOption: string;
}

declare module 'HighlightsWebPartStrings' {
    const strings: IHighlightsWebPartStrings;
    export = strings;
}
