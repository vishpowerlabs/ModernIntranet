/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

declare interface IQuickLinksWebPartStrings {
    PropertyPaneDescription: string;
    DataSourceGroupName: string;
    ColumnMappingGroupName: string;
    DisplaySettingsGroupName: string;

    SiteUrlFieldLabel: string;
    ListIdFieldLabel: string;

    TitleColumnFieldLabel: string;
    LinkColumnFieldLabel: string;
    IconColumnFieldLabel: string;
    PinnedColumnFieldLabel: string;

    ColumnsPerRowFieldLabel: string;
    OpenInNewTabFieldLabel: string;
    ShowTitleFieldLabel: string;
    TitleFieldLabel: string;
    ShowBackgroundBarFieldLabel: string;
}

declare module 'QuickLinksWebPartStrings' {
    const strings: IQuickLinksWebPartStrings;
    export = strings;
}
