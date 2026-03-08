/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneDropdown,
    PropertyPaneToggle,
    PropertyPaneTextField,
    PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CalendarWebPartStrings';
import Calendar from './components/Calendar';
import { ICalendarProps } from './components/ICalendarProps';

import { SiteListService } from '../../common/services/SiteListService';
import { ThemeService } from '../../common/services/ThemeService';
import {
    PropertyFieldSitePicker,
    PropertyFieldListPicker,
    PropertyFieldColumnPicker
} from '../../common/propertyPaneControls';

export interface ICalendarWebPartProps {
    siteUrl: string;
    listId: string;
    titleColumn: string;
    dateColumn: string;
    endDateColumn: string;
    locationColumn: string;
    defaultView: 'day' | 'week' | 'month' | 'year';
    showTitle: boolean;
    title: string;
    showBackgroundBar: boolean;
    titleBarStyle: 'solid' | 'underline';
    yearViewType: 'grid' | 'timeline';
}

export default class CalendarWebPart extends BaseClientSideWebPart<ICalendarWebPartProps> {
    private _siteListService!: SiteListService;

    protected async onInit(): Promise<void> {
        await super.onInit();
        this._siteListService = new SiteListService(this.context);
        ThemeService.initialize(this.context);

        if (!this.properties.siteUrl) {
            this.properties.siteUrl = this.context.pageContext.web.absoluteUrl;
        }

        // Default column mappings for standard SP Events (Calendar) list
        if (!this.properties.titleColumn) this.properties.titleColumn = 'Title';
        if (!this.properties.dateColumn) this.properties.dateColumn = 'EventDate';
        if (!this.properties.endDateColumn) this.properties.endDateColumn = 'EndDate';
        if (!this.properties.locationColumn) this.properties.locationColumn = 'Location';
    }

    public render(): void {
        const element: React.ReactElement<ICalendarProps> = React.createElement(
            Calendar,
            {
                context: this.context,
                siteUrl: this.properties.siteUrl || this.context.pageContext.web.absoluteUrl,
                listId: this.properties.listId,
                titleColumn: this.properties.titleColumn,
                dateColumn: this.properties.dateColumn,
                endDateColumn: this.properties.endDateColumn,
                locationColumn: this.properties.locationColumn,
                defaultView: this.properties.defaultView || 'month',
                showTitle: this.properties.showTitle,
                title: this.properties.title,
                showBackgroundBar: this.properties.showBackgroundBar,
                titleBarStyle: this.properties.titleBarStyle,
                yearViewType: this.properties.yearViewType || 'grid'
            }
        );

        this.domElement.style.cssText += ThemeService.getThemeCSS().split(';').join(' !important;');
        ReactDom.render(element, this.domElement);
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyFieldSitePicker('siteUrl', {
                                    label: strings.SiteUrlFieldLabel,
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'calendarSiteUrlPicker'
                                }),
                                PropertyFieldListPicker('listId', {
                                    label: strings.ListIdFieldLabel,
                                    siteUrl: this.properties.siteUrl,
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'calendarListIdPicker',
                                    disabled: !this.properties.siteUrl
                                })
                            ]
                        },
                        {
                            groupName: strings.ColumnsGroupName,
                            groupFields: [
                                this._getColumnPicker('titleColumn', strings.TitleColumnFieldLabel, 'Text', 'calendarTitleColumnPicker'),
                                this._getColumnPicker('dateColumn', strings.DateColumnFieldLabel, 'DateTime', 'calendarDateColumnPicker'),
                                this._getColumnPicker('endDateColumn', strings.EndDateColumnFieldLabel, 'DateTime', 'calendarEndDateColumnPicker'),
                                this._getColumnPicker('locationColumn', strings.LocationColumnFieldLabel, 'Text', 'calendarLocationColumnPicker')
                            ]
                        },
                        {
                            groupName: strings.DisplayGroupName,
                            groupFields: [
                                PropertyPaneDropdown('defaultView', {
                                    label: strings.DefaultViewFieldLabel,
                                    options: [
                                        { key: 'day', text: 'Day' },
                                        { key: 'week', text: 'Week' },
                                        { key: 'month', text: 'Month' },
                                        { key: 'year', text: 'Year' }
                                    ],
                                    selectedKey: 'month'
                                }),
                                PropertyPaneDropdown('yearViewType', {
                                    label: strings.YearViewTypeFieldLabel,
                                    options: [
                                        { key: 'grid', text: strings.YearViewTypeGrid },
                                        { key: 'timeline', text: strings.YearViewTypeTimeline }
                                    ],
                                    selectedKey: 'grid'
                                }),
                                PropertyPaneToggle('showTitle', {
                                    label: strings.ShowTitleFieldLabel
                                }),
                                PropertyPaneTextField('title', {
                                    label: strings.TitleFieldLabel,
                                    disabled: !this.properties.showTitle
                                }),
                                PropertyPaneToggle('showBackgroundBar', {
                                    label: strings.ShowBackgroundBarFieldLabel
                                }),
                                ...(this.properties.showBackgroundBar ? [
                                    PropertyPaneChoiceGroup('titleBarStyle', {
                                        label: strings.TitleBarStyleFieldLabel,
                                        options: [
                                            { key: 'solid', text: strings.TitleBarStyleSolidOption, iconProps: { officeFabricIconFontName: 'ChromeBack' } },
                                            { key: 'underline', text: strings.TitleBarStyleUnderlineOption, iconProps: { officeFabricIconFontName: 'ChromeMinimize' } }
                                        ]
                                    })
                                ] : [])
                            ]
                        }
                    ]
                }
            ]
        };
    }

    private _getColumnPicker(propertyPath: string, label: string, typeFilter: string, key: string): any {
        return PropertyFieldColumnPicker(propertyPath, {
            label,
            siteUrl: this.properties.siteUrl,
            listId: this.properties.listId,
            typeFilter: typeFilter as any,
            siteListService: this._siteListService,
            onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
            properties: this.properties,
            wpContext: this.context,
            key,
            disabled: !this.properties.listId
        });
    }
}
