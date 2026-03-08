import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneToggle,
    PropertyPaneTextField,
    PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'FaqWebPartStrings';
import Faq from './components/Faq';
import { IFaqProps } from './components/IFaqProps';

import { SiteListService } from '../../common/services/SiteListService';
import { ThemeService } from '../../common/services/ThemeService';
import {
    PropertyFieldSitePicker,
    PropertyFieldListPicker,
    PropertyFieldColumnPicker
} from '../../common/propertyPaneControls';

export interface IFaqWebPartProps {
    showTitle: boolean;
    title: string;
    showBackgroundBar: boolean;
    titleBarStyle: 'solid' | 'underline';

    siteUrl: string;
    listId: string;

    questionColumn: string;
    answerColumn: string;
    categoryColumn: string;
    orderColumn: string;

    showSearch: boolean;
    showCategoryFilter: boolean;
    allowMultipleOpen: boolean;
    expandFirstItem: boolean;
}

export default class FaqWebPart extends BaseClientSideWebPart<IFaqWebPartProps> {
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
        const element: React.ReactElement<IFaqProps> = React.createElement(
            Faq,
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
        return {
            pages: [
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
                                PropertyFieldSitePicker('siteUrl', {
                                    label: strings.SiteUrlFieldLabel,
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'faqSitePicker'
                                }),
                                PropertyFieldListPicker('listId', {
                                    label: strings.ListIdFieldLabel,
                                    siteUrl: this.properties.siteUrl,
                                    siteListService: this._siteListService,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    wpContext: this.context,
                                    key: 'faqListPicker',
                                    disabled: !this.properties.siteUrl
                                })
                            ]
                        },
                        {
                            groupName: strings.ColumnMappingsGroupName,
                            groupFields: [
                                this._getColumnPicker('questionColumn', strings.QuestionColumnFieldLabel, 'Text', 'faqQCol'),
                                this._getColumnPicker('answerColumn', strings.AnswerColumnFieldLabel, 'Note', 'faqACol'),
                                this._getColumnPicker('categoryColumn', strings.CategoryColumnFieldLabel, 'Choice,Text', 'faqCCol'),
                                this._getColumnPicker('orderColumn', strings.OrderColumnFieldLabel, 'Number,Text', 'faqOCol')
                            ]
                        },
                        {
                            groupName: strings.DisplaySettingsGroupName,
                            groupFields: [
                                PropertyPaneToggle('showSearch', { label: strings.ShowSearchFieldLabel }),
                                PropertyPaneToggle('showCategoryFilter', { label: strings.ShowCategoryFilterFieldLabel }),
                                PropertyPaneToggle('allowMultipleOpen', { label: strings.AllowMultipleOpenFieldLabel }),
                                PropertyPaneToggle('expandFirstItem', { label: strings.ExpandFirstItemFieldLabel })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
