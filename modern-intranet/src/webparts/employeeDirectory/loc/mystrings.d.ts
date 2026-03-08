declare interface IEmployeeDirectoryWebPartStrings {
    PropertyPaneDescription: string;
    DataSourceGroupName: string;
    ColumnMappingsGroupName: string;
    DisplaySettingsGroupName: string;

    SourceFieldLabel: string;
    SourceGraphOption: string;
    SourceSPListOption: string;

    SiteUrlFieldLabel: string;
    ListIdFieldLabel: string;

    NameColumnFieldLabel: string;
    PhotoColumnFieldLabel: string;
    JobTitleColumnFieldLabel: string;
    DepartmentColumnFieldLabel: string;
    LocationColumnFieldLabel: string;
    EmailColumnFieldLabel: string;
    PhoneColumnFieldLabel: string;
    ManagerColumnFieldLabel: string;
    ProjectsColumnFieldLabel: string;
    AboutMeColumnFieldLabel: string;
    InterestsColumnFieldLabel: string;
    SkillsColumnFieldLabel: string;

    ViewModeFieldLabel: string;
    ViewModeListOption: string;
    ViewModeGridOption: string;

    PageSizeFieldLabel: string;
    ShowFiltersFieldLabel: string;
    ShowPaginationFieldLabel: string;

    BasicGroupName: string;
    ShowTitleFieldLabel: string;
    TitleFieldLabel: string;
    ShowBackgroundBarFieldLabel: string;
    TitleBarStyleFieldLabel: string;
    TitleBarStyleSolidOption: string;
    TitleBarStyleUnderlineOption: string;
}

declare module 'EmployeeDirectoryWebPartStrings' {
    const strings: IEmployeeDirectoryWebPartStrings;
    export = strings;
}
