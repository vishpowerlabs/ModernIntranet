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
    PropertyPaneDropdown,
    PropertyPaneToggle,
    PropertyPaneTextField,
    PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'HighlightsWebPartStrings';
import { Highlights } from './components/Highlights';
import { IHighlightsProps } from './components/IHighlightsProps';
import { SiteListService } from '../../common/services/SiteListService';
import { ThemeService } from '../../common/services/ThemeService';
import { PropertyFieldSitePicker } from '../../common/propertyPaneControls/PropertyFieldSitePicker';
import { PropertyFieldListPicker } from '../../common/propertyPaneControls/PropertyFieldListPicker';
import { PropertyFieldColumnPicker } from '../../common/propertyPaneControls/PropertyFieldColumnPicker';

export interface IHighlightsWebPartProps {
    siteUrl: string;
    listId: string;
    titleColumn: string;
    descriptionColumn: string;
    bannerImageColumn: string;
    linkColumn: string;
    pinnedColumn: string;
    maxItems: number;
    columns: number;
    showTitle: boolean;
    title: string;
    showBackgroundBar: boolean;
    titleBarStyle: 'solid' | 'underline';
}

export default class HighlightsWebPart extends BaseClientSideWebPart<IHighlightsWebPartProps> {

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
        const element: React.ReactElement<IHighlightsProps> = React.createElement(
            Highlights,
            {
                siteUrl: this.properties.siteUrl,
                listId: this.properties.listId,
                titleColumn: this.properties.titleColumn,
                descriptionColumn: this.properties.descriptionColumn,
                bannerImageColumn: this.properties.bannerImageColumn,
                linkColumn: this.properties.linkColumn,
                pinnedColumn: this.properties.pinnedColumn,
                maxItems: this.properties.maxItems,
                columns: this.properties.columns || 3,
                showTitle: this.properties.showTitle,
                title: this.properties.title,
                showBackgroundBar: this.properties.showBackgroundBar,
                titleBarStyle: this.properties.titleBarStyle || 'underline',
                siteId: this.context.pageContext.site.id.toString(),
                webId: this.context.pageContext.web.id.toString(),
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
                            groupName: strings.DataSourceGroupName,
                            groupFields: [
                                PropertyFieldSitePicker('siteUrl', {
                                    label: strings.SiteUrlFieldLabel,
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'siteUrlPicker'
                                }),
                                PropertyFieldListPicker('listId', {
                                    label: strings.ListIdFieldLabel,
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
                            groupName: strings.ColumnMappingGroupName,
                            groupFields: [
                                PropertyFieldColumnPicker('titleColumn', {
                                    label: strings.TitleColumnFieldLabel,
                                    siteUrl: this.properties.siteUrl,
                                    listId: this.properties.listId,
                                    typeFilter: 'Text',
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'titleColumnPicker',
                                    disabled: !this.properties.listId
                                }),
                                PropertyFieldColumnPicker('descriptionColumn', {
                                    label: strings.DescriptionColumnFieldLabel,
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
                                PropertyFieldColumnPicker('bannerImageColumn', {
                                    label: strings.BannerImageColumnFieldLabel,
                                    siteUrl: this.properties.siteUrl,
                                    listId: this.properties.listId,
                                    typeFilter: 'Thumbnail,URL,Image',
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'bannerImageColumnPicker',
                                    disabled: !this.properties.listId
                                }),
                                PropertyFieldColumnPicker('linkColumn', {
                                    label: strings.LinkColumnFieldLabel,
                                    siteUrl: this.properties.siteUrl,
                                    listId: this.properties.listId,
                                    typeFilter: 'URL',
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'linkColumnPicker',
                                    disabled: !this.properties.listId
                                }),
                                PropertyFieldColumnPicker('pinnedColumn', {
                                    label: strings.PinnedColumnFieldLabel,
                                    siteUrl: this.properties.siteUrl,
                                    listId: this.properties.listId,
                                    typeFilter: 'Boolean,Choice,Bit',
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'pinnedColumnPicker',
                                    disabled: !this.properties.listId
                                })
                            ]
                        },
                        {
                            groupName: strings.DisplaySettingsGroupName,
                            groupFields: [
                                PropertyPaneSlider('maxItems', {
                                    label: strings.MaxItemsFieldLabel,
                                    min: 3,
                                    max: 12,
                                    step: 1
                                }),
                                PropertyPaneDropdown('columns', {
                                    label: strings.ColumnsFieldLabel,
                                    options: [
                                        { key: 2, text: '2 Columns' },
                                        { key: 3, text: '3 Columns' }
                                    ]
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
}
