/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

declare interface IEventsWebPartStrings {
  PropertyPaneDescription: string;
  DataSourceGroupName: string;
  ColumnMappingGroupName: string;
  DisplaySettingsGroupName: string;
  SiteUrlFieldLabel: string;
  ListIdFieldLabel: string;
  TitleColumnFieldLabel: string;
  DateColumnFieldLabel: string;
  ActiveColumnFieldLabel: string;
  ImageColumnFieldLabel: string;
  LinkColumnFieldLabel: string;
  LocationColumnFieldLabel: string;
  DescriptionColumnFieldLabel: string;
  PinnedColumnFieldLabel: string;
  MaxItemsFieldLabel: string;
  ItemsPerRowFieldLabel: string;
  ShowViewAllFieldLabel: string;
  ViewAllUrlFieldLabel: string;
  ShowTitleFieldLabel: string;
  TitleFieldLabel: string;
  ShowBackgroundBarFieldLabel: string;
  TitleBarStyleFieldLabel: string;
  TitleBarStyleSolidOption: string;
  TitleBarStyleUnderlineOption: string;
  ShowPaginationFieldLabel: string;
}

declare module 'EventsWebPartStrings' {
  const strings: IEventsWebPartStrings;
  export = strings;
}
