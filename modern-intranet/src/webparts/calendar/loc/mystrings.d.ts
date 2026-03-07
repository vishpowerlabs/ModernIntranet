declare interface ICalendarWebPartStrings {
    PropertyPaneDescription: string;
    BasicGroupName: string;
    SiteUrlFieldLabel: string;
    ListIdFieldLabel: string;
    ColumnsGroupName: string;
    TitleColumnFieldLabel: string;
    DateColumnFieldLabel: string;
    EndDateColumnFieldLabel: string;
    LocationColumnFieldLabel: string;
    DisplayGroupName: string;
    DefaultViewFieldLabel: string;
    YearViewTypeFieldLabel: string;
    YearViewTypeGrid: string;
    YearViewTypeTimeline: string;
    ShowTitleFieldLabel: string;
    TitleFieldLabel: string;
    ShowBackgroundBarFieldLabel: string;
}

declare module 'CalendarWebPartStrings' {
    const strings: ICalendarWebPartStrings;
    export = strings;
}
