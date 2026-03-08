import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneDropdown,
    PropertyPaneToggle,
    PropertyPaneTextField,
    PropertyPaneSlider,
    PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'EmployeeDirectoryWebPartStrings';
import EmployeeDirectory from './components/EmployeeDirectory';
import { IEmployeeDirectoryProps } from './components/IEmployeeDirectoryProps';

import { SiteListService } from '../../common/services/SiteListService';
import { ThemeService } from '../../common/services/ThemeService';
import {
    PropertyFieldSitePicker,
    PropertyFieldListPicker,
    PropertyFieldColumnPicker
} from '../../common/propertyPaneControls';

export interface IEmployeeDirectoryWebPartProps {
    showTitle: boolean;
    title: string;
    showBackgroundBar: boolean;
    titleBarStyle: 'solid' | 'underline';

    source: 'graph' | 'spList';
    siteUrl: string;
    listId: string;

    nameColumn: string;
    photoColumn: string;
    jobTitleColumn: string;
    departmentColumn: string;
    locationColumn: string;
    emailColumn: string;
    phoneColumn: string;
    managerColumn: string;
    projectsColumn: string;
    aboutMeColumn: string;
    interestsColumn: string;
    skillsColumn: string;

    viewMode: 'list' | 'grid';
    pageSize: number;
    showFilters: boolean;
    showPagination: boolean;
}

export default class EmployeeDirectoryWebPart extends BaseClientSideWebPart<IEmployeeDirectoryWebPartProps> {
    private _siteListService!: SiteListService;

    protected async onInit(): Promise<void> {
        await super.onInit();
        this._siteListService = new SiteListService(this.context);
        ThemeService.initialize(this.context);

        if (!this.properties.siteUrl) {
            this.properties.siteUrl = this.context.pageContext.web.absoluteUrl;
        }
    }

    public render(): void {
        const element: React.ReactElement<IEmployeeDirectoryProps> = React.createElement(
            EmployeeDirectory,
            {
                context: this.context,
                ...this.properties,
                siteUrl: this.properties.siteUrl || this.context.pageContext.web.absoluteUrl,
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

    private _getColumnPicker(propertyPath: string, label: string, typeFilter: string, key: string, disabled: boolean = false): any {
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
            disabled: disabled || !this.properties.listId
        });
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        const isSpList = this.properties.source === 'spList';

        const columnMappingsGroup = {
            groupName: strings.ColumnMappingsGroupName,
            groupFields: [
                this._getColumnPicker('nameColumn', strings.NameColumnFieldLabel, 'Text', 'dirNameCol'),
                this._getColumnPicker('photoColumn', strings.PhotoColumnFieldLabel, 'Thumbnail,URL,Image', 'dirPhotoCol'),
                this._getColumnPicker('jobTitleColumn', strings.JobTitleColumnFieldLabel, 'Text', 'dirJobCol'),
                this._getColumnPicker('departmentColumn', strings.DepartmentColumnFieldLabel, 'Text', 'dirDeptCol'),
                this._getColumnPicker('locationColumn', strings.LocationColumnFieldLabel, 'Text', 'dirLocCol'),
                this._getColumnPicker('emailColumn', strings.EmailColumnFieldLabel, 'User', 'dirEmailCol'),
                this._getColumnPicker('phoneColumn', strings.PhoneColumnFieldLabel, 'Text', 'dirPhoneCol'),
                this._getColumnPicker('managerColumn', strings.ManagerColumnFieldLabel, 'User', 'dirMgrCol'),
                this._getColumnPicker('projectsColumn', strings.ProjectsColumnFieldLabel, 'Note', 'dirProjCol'),
                this._getColumnPicker('aboutMeColumn', strings.AboutMeColumnFieldLabel, 'Note', 'dirAboutCol'),
                this._getColumnPicker('interestsColumn', strings.InterestsColumnFieldLabel, 'Note', 'dirIntCol'),
                this._getColumnPicker('skillsColumn', strings.SkillsColumnFieldLabel, 'Note', 'dirSkillsCol')
            ]
        };

        const pages = [
            {
                header: { description: strings.PropertyPaneDescription },
                groups: [
                    {
                        groupName: strings.BasicGroupName,
                        groupFields: [
                            PropertyPaneToggle('showTitle', { label: strings.ShowTitleFieldLabel }),
                            PropertyPaneTextField('title', { label: strings.TitleFieldLabel, disabled: !this.properties.showTitle }),
                            PropertyPaneToggle('showBackgroundBar', { label: strings.ShowBackgroundBarFieldLabel }),
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
                    },
                    {
                        groupName: strings.DataSourceGroupName,
                        groupFields: [
                            PropertyPaneDropdown('source', {
                                label: strings.SourceFieldLabel,
                                options: [
                                    { key: 'graph', text: strings.SourceGraphOption },
                                    { key: 'spList', text: strings.SourceSPListOption }
                                ]
                            }),
                            ...(isSpList ? [
                                PropertyFieldSitePicker('siteUrl', {
                                    label: strings.SiteUrlFieldLabel,
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'dirSitePicker'
                                }),
                                PropertyFieldListPicker('listId', {
                                    label: strings.ListIdFieldLabel,
                                    siteUrl: this.properties.siteUrl,
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'dirListPicker',
                                    disabled: !this.properties.siteUrl
                                })
                            ] : [])
                        ]
                    }
                ]
            }
        ];

        if (isSpList) {
            pages[0].groups.push(columnMappingsGroup as any);
        }

        pages[0].groups.push({
            groupName: strings.DisplaySettingsGroupName,
            groupFields: [
                PropertyPaneDropdown('viewMode', {
                    label: strings.ViewModeFieldLabel,
                    options: [
                        { key: 'list', text: strings.ViewModeListOption },
                        { key: 'grid', text: strings.ViewModeGridOption }
                    ]
                }),
                PropertyPaneSlider('pageSize', {
                    label: strings.PageSizeFieldLabel,
                    min: 5,
                    max: 25,
                    step: 1,
                    showValue: true
                }),
                PropertyPaneToggle('showFilters', { label: strings.ShowFiltersFieldLabel }),
                PropertyPaneToggle('showPagination', { label: strings.ShowPaginationFieldLabel })
            ]
        });

        return { pages };
    }
}
