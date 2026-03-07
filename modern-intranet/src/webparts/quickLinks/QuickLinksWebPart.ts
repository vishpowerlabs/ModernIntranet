import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneDropdown,
    PropertyPaneToggle,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'QuickLinksWebPartStrings';

import { QuickLinks } from './components/QuickLinks';
import { IQuickLinksProps } from './components/IQuickLinksProps';

import { SiteListService } from '../../common/services/SiteListService';
import { ThemeService } from '../../common/services/ThemeService';
import {
    PropertyFieldSitePicker,
    PropertyFieldListPicker,
    PropertyFieldColumnPicker
} from '../../common/propertyPaneControls';

export interface IQuickLinksWebPartProps {
    siteUrl: string;
    listId: string;
    titleColumn: string;
    linkColumn: string;
    iconColumn: string;
    pinnedColumn: string;
    columnsPerRow: number;
    openInNewTab: boolean;
    showTitle: boolean;
    title: string;
    showBackgroundBar: boolean;
}

export default class QuickLinksWebPart extends BaseClientSideWebPart<IQuickLinksWebPartProps> {

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
        const element: React.ReactElement<IQuickLinksProps> = React.createElement(
            QuickLinks,
            {
                siteUrl: this.properties.siteUrl,
                listId: this.properties.listId,
                titleColumn: this.properties.titleColumn,
                linkColumn: this.properties.linkColumn,
                iconColumn: this.properties.iconColumn,
                pinnedColumn: this.properties.pinnedColumn,
                columnsPerRow: this.properties.columnsPerRow,
                openInNewTab: this.properties.openInNewTab,
                showTitle: this.properties.showTitle,
                title: this.properties.title,
                showBackgroundBar: this.properties.showBackgroundBar,
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
                                PropertyFieldColumnPicker('iconColumn', {
                                    label: strings.IconColumnFieldLabel,
                                    siteUrl: this.properties.siteUrl,
                                    listId: this.properties.listId,
                                    typeFilter: 'Text',
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'iconColumnPicker',
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
                                PropertyPaneDropdown('columnsPerRow', {
                                    label: strings.ColumnsPerRowFieldLabel,
                                    options: [
                                        { key: 2, text: '2' },
                                        { key: 3, text: '3' },
                                        { key: 4, text: '4' },
                                        { key: 6, text: '6' }
                                    ]
                                }),
                                PropertyPaneToggle('openInNewTab', {
                                    label: strings.OpenInNewTabFieldLabel
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
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
