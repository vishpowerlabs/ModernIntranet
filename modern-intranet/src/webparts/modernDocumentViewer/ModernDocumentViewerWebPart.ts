import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneSlider,
    PropertyPaneDropdown,
    PropertyPaneTextField,
    PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'ModernDocumentViewerWebPartStrings';
import { ModernDocumentViewer } from './components/ModernDocumentViewer';
import { IModernDocumentViewerProps } from './components/IModernDocumentViewerProps';
import { SiteListService } from '../../common/services/SiteListService';
import { ThemeService } from '../../common/services/ThemeService';
import {
    PropertyFieldSitePicker,
    PropertyFieldListPicker,
    PropertyFieldColumnPicker
} from '../../common/propertyPaneControls';

export interface IModernDocumentViewerWebPartProps {
    siteUrl: string;
    listId: string;
    categoryField: string;
    subCategoryField: string;
    descriptionField: string;
    pinnedField: string;
    enableSubCategory: boolean;
    categoryDisplayType: 'side' | 'top';
    pageSize: number;
    webPartTitle: string;
    webPartTitleFontSize: string;
    webPartDescription: string;
    webPartDescriptionFontSize: string;
    headerOpacity: number;
}

export default class ModernDocumentViewerWebPart extends BaseClientSideWebPart<IModernDocumentViewerWebPartProps> {

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
        const element: React.ReactElement<IModernDocumentViewerProps> = React.createElement(
            ModernDocumentViewer,
            {
                siteUrl: this.properties.siteUrl,
                listId: this.properties.listId,
                categoryField: this.properties.categoryField || 'Category',
                subCategoryField: this.properties.subCategoryField || 'SubCategory',
                descriptionField: this.properties.descriptionField || 'Description',
                pinnedField: this.properties.pinnedField,
                enableSubCategory: this.properties.enableSubCategory ?? true,
                categoryDisplayType: this.properties.categoryDisplayType || 'side',
                pageSize: this.properties.pageSize || 10,
                webPartTitle: this.properties.webPartTitle,
                webPartTitleFontSize: this.properties.webPartTitleFontSize || '24px',
                webPartDescription: this.properties.webPartDescription,
                webPartDescriptionFontSize: this.properties.webPartDescriptionFontSize || '14px',
                headerOpacity: this.properties.headerOpacity ?? 1,
                context: this.context
            }
        );

        // Apply exact theme variables to the web part container for child scoping
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
                            groupName: "Data Source",
                            groupFields: [
                                PropertyFieldSitePicker('siteUrl', {
                                    label: 'Select Site',
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'siteUrlPicker'
                                }),
                                PropertyFieldListPicker('listId', {
                                    label: 'Select Document Library',
                                    siteUrl: this.properties.siteUrl,
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'listIdPicker',
                                    disabled: !this.properties.siteUrl
                                })
                            ]
                        },
                        {
                            groupName: "Column Mapping",
                            groupFields: [
                                PropertyFieldColumnPicker('categoryField', {
                                    label: 'Category Column',
                                    siteUrl: this.properties.siteUrl,
                                    listId: this.properties.listId,
                                    typeFilter: 'Text,Choice',
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'categoryColumnPicker',
                                    disabled: !this.properties.listId
                                }),
                                PropertyFieldColumnPicker('subCategoryField', {
                                    label: 'Sub-Category Column',
                                    siteUrl: this.properties.siteUrl,
                                    listId: this.properties.listId,
                                    typeFilter: 'Text,Choice',
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'subCategoryColumnPicker',
                                    disabled: !this.properties.listId
                                }),
                                PropertyFieldColumnPicker('descriptionField', {
                                    label: 'Description Column',
                                    siteUrl: this.properties.siteUrl,
                                    listId: this.properties.listId,
                                    typeFilter: 'Text,Note',
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'descriptionColumnPicker',
                                    disabled: !this.properties.listId
                                }),
                                PropertyFieldColumnPicker('pinnedField', {
                                    label: 'Pinned Column',
                                    siteUrl: this.properties.siteUrl,
                                    listId: this.properties.listId,
                                    typeFilter: 'Boolean,Choice,Bit',
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'pinnedColumnPicker',
                                    disabled: !this.properties.listId
                                }),
                                PropertyPaneToggle('enableSubCategory', {
                                    label: 'Enable Sub-Categories',
                                    onText: 'On',
                                    offText: 'Off'
                                }),
                                PropertyPaneDropdown('categoryDisplayType', {
                                    label: 'Category Display Style',
                                    options: [
                                        { key: 'side', text: 'Left Navigation' },
                                        { key: 'top', text: 'Top Tabs' }
                                    ],
                                    disabled: this.properties.enableSubCategory === true
                                })
                            ]
                        },
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('webPartTitle', {
                                    label: 'Web Part Title'
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
                                PropertyPaneTextField('webPartDescription', {
                                    label: 'Web Part Description',
                                    multiline: true
                                }),
                                PropertyPaneDropdown('webPartDescriptionFontSize', {
                                    label: 'Description Font Size',
                                    options: [
                                        { key: '12px', text: 'Small (12px)' },
                                        { key: '14px', text: 'Medium (14px)' },
                                        { key: '16px', text: 'Large (16px)' },
                                        { key: '18px', text: 'Extra Large (18px)' }
                                    ]
                                }),
                                PropertyPaneSlider('headerOpacity', {
                                    label: 'Header Opacity',
                                    min: 0,
                                    max: 1,
                                    step: 0.1,
                                    showValue: true
                                }),
                                PropertyPaneTextField('pageSize', {
                                    label: 'Items per page'
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
