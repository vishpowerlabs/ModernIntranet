/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneSlider,
    PropertyPaneTextField,
    PropertyPaneToggle,
    PropertyPaneChoiceGroup,
    PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'EmployeeSpotlightWebPartStrings';
import EmployeeSpotlight from './components/EmployeeSpotlight';
import { IEmployeeSpotlightProps } from './components/IEmployeeSpotlightProps';
import { SiteListService } from '../../common/services/SiteListService';
import { ThemeService } from '../../common/services/ThemeService';
import {
    PropertyFieldSitePicker,
    PropertyFieldListPicker,
    PropertyFieldColumnPicker,
    PropertyFieldPeoplePicker
} from '../../common/propertyPaneControls';

export interface IEmployeeSpotlightWebPartProps {
    siteUrl: string;
    listId: string;

    nameColumn: string;
    photoColumn: string;
    jobTitleColumn: string;
    departmentColumn: string;
    emailColumn: string;
    spotlightColumn: string;
    spotlightTextColumn: string;

    maxItems: number;
    autoRotateInterval: number;

    showTitle: boolean;
    title: string;
    webPartTitleFontSize: string;
    showBackgroundBar: boolean;
    titleBarStyle: 'solid' | 'underline';
    layoutMode: 'standard' | 'compact';

    source: 'spList' | 'graph';
    selectedUsers: any[];
    commonDescription: string;
}

export default class EmployeeSpotlightWebPart extends BaseClientSideWebPart<IEmployeeSpotlightWebPartProps> {

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
        const element: React.ReactElement<IEmployeeSpotlightProps> = React.createElement(
            EmployeeSpotlight,
            {
                ...this.properties,
                siteUrl: this.properties.siteUrl,
                listId: this.properties.listId,
                nameColumn: this.properties.nameColumn,
                photoColumn: this.properties.photoColumn,
                jobTitleColumn: this.properties.jobTitleColumn,
                departmentColumn: this.properties.departmentColumn,
                emailColumn: this.properties.emailColumn,
                spotlightColumn: this.properties.spotlightColumn,
                spotlightTextColumn: this.properties.spotlightTextColumn,
                maxItems: this.properties.maxItems,
                autoRotateInterval: this.properties.autoRotateInterval,
                showTitle: this.properties.showTitle,
                title: this.properties.title,
                webPartTitleFontSize: this.properties.webPartTitleFontSize || '24px',
                showBackgroundBar: this.properties.showBackgroundBar ?? true,
                titleBarStyle: this.properties.titleBarStyle || 'underline',
                layoutMode: this.properties.layoutMode || 'standard',
                source: this.properties.source || 'spList',
                selectedUsers: this.properties.selectedUsers || [],
                commonDescription: this.properties.commonDescription || '',
                context: this.context
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

    private _getColumnPicker(targetProperty: string, label: string, typeFilter: string, key: string, disabled: boolean = false): any {
        return PropertyFieldColumnPicker(targetProperty, {
            label: label,
            siteUrl: this.properties.siteUrl,
            listId: this.properties.listId,
            typeFilter: typeFilter,
            siteListService: this._siteListService,
            onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
            properties: this.properties,
            wpContext: this.context,
            key: key,
            disabled: !this.properties.listId || disabled
        });
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
                            groupName: strings.DataSourceGroupName,
                            groupFields: [
                                PropertyFieldSitePicker('siteUrl', {
                                    label: 'Select Site',
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'spotlightSiteUrlPicker'
                                }),
                                PropertyFieldListPicker('listId', {
                                    label: 'Choose Employee List',
                                    siteUrl: this.properties.siteUrl,
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'spotlightListIdPicker',
                                    disabled: !this.properties.siteUrl
                                })
                            ]
                        },
                        {
                            groupName: 'Data Source Selection',
                            groupFields: [
                                PropertyPaneChoiceGroup('source', {
                                    label: 'Select Data Source',
                                    options: [
                                        { key: 'spList', text: 'SharePoint List', iconProps: { officeFabricIconFontName: 'List' } },
                                        { key: 'graph', text: 'Microsoft Graph (Manual)', iconProps: { officeFabricIconFontName: 'People' } }
                                    ]
                                })
                            ]
                        },
                        {
                            groupName: strings.ColumnMappingGroupName,
                            groupFields: this.properties.source === 'graph' ? [
                                PropertyFieldPeoplePicker('selectedUsers', {
                                    label: 'Choose Employees',
                                    wpContext: this.context,
                                    properties: this.properties,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    key: 'peoplePicker',
                                    itemLimit: this.properties.maxItems || 10
                                }),
                                PropertyPaneTextField('commonDescription', {
                                    label: 'Common Spotlight Description',
                                    multiline: true,
                                    rows: 3
                                })
                            ] : [
                                this._getColumnPicker('nameColumn', 'Name Column', 'Text', 'namePicker'),
                                this._getColumnPicker('photoColumn', 'Photo Column (Optional)', 'Thumbnail,URL,Image', 'photoPicker'),
                                this._getColumnPicker('jobTitleColumn', 'Job Title Column', 'Text', 'jobTitlePicker'),
                                this._getColumnPicker('departmentColumn', 'Department Column', 'Text', 'deptPicker'),
                                this._getColumnPicker('emailColumn', 'Email/Person Column', 'User', 'emailPicker'),
                                this._getColumnPicker('spotlightColumn', 'Spotlight Flag Column', 'Boolean', 'spotlightPicker'),
                                this._getColumnPicker('spotlightTextColumn', 'Spotlight Write-up Column', 'Text,Note', 'spotlightTextPicker')
                            ]
                        },

                        {
                            groupName: strings.DisplayGroupName,
                            groupFields: [
                                PropertyPaneSlider('maxItems', {
                                    label: 'Maximum Spotlights to display',
                                    min: 1,
                                    max: 10,
                                    step: 1,
                                    value: 3
                                }),
                                PropertyPaneSlider('autoRotateInterval', {
                                    label: 'Auto-rotate seconds',
                                    min: 3,
                                    max: 15,
                                    step: 1,
                                    value: 5
                                })
                            ]
                        },
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneToggle('showTitle', { label: strings.ShowTitleFieldLabel }),
                                PropertyPaneTextField('title', {
                                    label: strings.TitleFieldLabel,
                                    disabled: !this.properties.showTitle
                                }),
                                PropertyPaneDropdown('webPartTitleFontSize', {
                                    label: 'Title Font Size',
                                    options: [
                                        { key: '16px', text: 'Small (16px)' },
                                        { key: '20px', text: 'Medium (20px)' },
                                        { key: '24px', text: 'Large (24px)' },
                                        { key: '32px', text: 'Extra Large (32px)' }
                                    ]
                                }),
                                PropertyPaneChoiceGroup('layoutMode', {
                                    label: 'Layout Mode',
                                    options: [
                                        { key: 'standard', text: 'Standard (Wide)', iconProps: { officeFabricIconFontName: 'WebAppBuilderFragment' } },
                                        { key: 'compact', text: 'Compact (Side Column)', iconProps: { officeFabricIconFontName: 'WebAppBuilderSlot' } }
                                    ]
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
}
