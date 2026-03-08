declare interface ICustomFaqWebPartStrings {
  PropertyPaneDescription: string;
  DisplaySettingsGroupName: string;
  DataSourceGroupName: string;
  AccordionBehaviorGroupName: string;
  WebpartTitleLabel: string;
  WebpartDescriptionLabel: string;
  SelectListLabel: string;
  TitleColumnLabel: string;
  DescriptionColumnLabel: string;
  AllowMultipleExpandedLabel: string;
}

declare module 'CustomFaqWebPartStrings' {
  const strings: ICustomFaqWebPartStrings;
  export = strings;
}
//vishnu to fix security issue based on scann