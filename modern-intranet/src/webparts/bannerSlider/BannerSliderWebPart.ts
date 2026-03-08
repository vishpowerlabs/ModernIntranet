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
    PropertyPaneToggle,
    PropertyPaneTextField,
    PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'BannerSliderWebPartStrings';
import { BannerSlider } from './components/BannerSlider';
import { IBannerSliderProps } from './components/IBannerSliderProps';

import { SiteListService } from '../../common/services/SiteListService';
import { ThemeService } from '../../common/services/ThemeService';
import { PropertyFieldSitePicker } from '../../common/propertyPaneControls/PropertyFieldSitePicker';
import { PropertyFieldListPicker } from '../../common/propertyPaneControls/PropertyFieldListPicker';
import { PropertyFieldColumnPicker } from '../../common/propertyPaneControls/PropertyFieldColumnPicker';

export interface IBannerSliderWebPartProps {
    siteUrl: string;
    listId: string;
    titleColumn: string;
    descriptionColumn: string;
    imageColumn: string;
    activeColumn: string;
    buttonTextColumn: string;
    pageLinkColumn: string;
    autoRotateInterval: number;
    showCta: boolean;
    showTitle: boolean;
    title: string;
    showBackgroundBar: boolean;
    titleBarStyle: 'solid' | 'underline';
}

export default class BannerSliderWebPart extends BaseClientSideWebPart<IBannerSliderWebPartProps> {

    private _siteListService!: SiteListService;
    protected async onInit(): Promise<void> {
        await super.onInit();
        this._siteListService = new SiteListService(this.context);
        ThemeService.initialize(this.context);

        // Default siteUrl to current site if not set
        if (!this.properties.siteUrl) {
            this.properties.siteUrl = this.context.pageContext.web.absoluteUrl;
        }
    }

    public render(): void {
        const element: React.ReactElement<IBannerSliderProps> = React.createElement(
            BannerSlider,
            {
                siteUrl: this.properties.siteUrl || this.context.pageContext.web.absoluteUrl,
                siteId: this.context.pageContext.site.id.toString(),
                webId: this.context.pageContext.web.id.toString(),
                listId: this.properties.listId,
                titleColumn: this.properties.titleColumn,
                descriptionColumn: this.properties.descriptionColumn,
                imageColumn: this.properties.imageColumn,
                activeColumn: this.properties.activeColumn,
                buttonTextColumn: this.properties.buttonTextColumn,
                pageLinkColumn: this.properties.pageLinkColumn,
                autoRotateInterval: this.properties.autoRotateInterval,
                showCta: this.properties.showCta,
                showTitle: this.properties.showTitle,
                title: this.properties.title,
                showBackgroundBar: this.properties.showBackgroundBar,
                titleBarStyle: this.properties.titleBarStyle || 'underline',
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
                                    typeFilter: 'Text,Note', // allow both single and multiline text
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'descriptionColumnPicker',
                                    disabled: !this.properties.listId
                                }),
                                PropertyFieldColumnPicker('imageColumn', {
                                    label: strings.ImageColumnFieldLabel,
                                    siteUrl: this.properties.siteUrl,
                                    listId: this.properties.listId,
                                    typeFilter: 'Thumbnail,URL,Image', // Image type
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'imageColumnPicker',
                                    disabled: !this.properties.listId
                                }),
                                PropertyFieldColumnPicker('activeColumn', {
                                    label: strings.ActiveColumnFieldLabel,
                                    siteUrl: this.properties.siteUrl,
                                    listId: this.properties.listId,
                                    typeFilter: 'Boolean',
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'activeColumnPicker',
                                    disabled: !this.properties.listId
                                }),
                                PropertyFieldColumnPicker('buttonTextColumn', {
                                    label: strings.ButtonTextColumnFieldLabel,
                                    siteUrl: this.properties.siteUrl,
                                    listId: this.properties.listId,
                                    typeFilter: 'Text',
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'buttonTextColumnPicker',
                                    disabled: !this.properties.listId
                                }),
                                PropertyFieldColumnPicker('pageLinkColumn', {
                                    label: strings.PageLinkColumnFieldLabel,
                                    siteUrl: this.properties.siteUrl,
                                    listId: this.properties.listId,
                                    typeFilter: 'URL',
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'pageLinkColumnPicker',
                                    disabled: !this.properties.listId
                                })
                            ]
                        },
                        {
                            groupName: strings.DisplaySettingsGroupName,
                            groupFields: [
                                PropertyPaneSlider('autoRotateInterval', {
                                    label: strings.AutoRotateIntervalFieldLabel,
                                    min: 3,
                                    max: 10,
                                    step: 1
                                }),
                                PropertyPaneToggle('showCta', {
                                    label: strings.ShowCtaFieldLabel
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
