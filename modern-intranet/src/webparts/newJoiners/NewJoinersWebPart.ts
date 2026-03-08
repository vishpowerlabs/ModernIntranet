/**
 * DEVELOPER BY VISHPOWERLABS
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
    PropertyPaneDropdown,
    PropertyPaneButton,
    PropertyPaneButtonType,
    PropertyPaneHorizontalRule
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'NewJoinersWebPartStrings';
import { NewJoiners } from './components/NewJoiners';
import { INewJoinersProps } from './components/INewJoinersProps';
import { SiteListService } from '../../common/services/SiteListService';
import { ThemeService } from '../../common/services/ThemeService';
import {
    PropertyFieldSitePicker,
    PropertyFieldListPicker,
    PropertyFieldColumnPicker,
    PropertyFieldPeoplePicker
} from '../../common/propertyPaneControls';

export interface INewJoinersWebPartProps {
    siteUrl: string;
    listId: string;
    nameColumn: string;
    photoColumn: string;
    jobTitleColumn: string;
    departmentColumn: string;
    emailColumn: string;
    newJoinerColumn: string;
    newJoinerTextColumn: string;
    maxItems: number;
    layout: 'list' | 'grid' | 'strip';
    layoutMode: 'standard' | 'compact';
    autoRotateInterval: number;
    source: 'spList' | 'graph';
    manualJoiners: any[];
    commonIntro: string;
    webPartTitle: string;
    showBackgroundBar: boolean;
    titleBarStyle: 'solid' | 'underline';
}

export default class NewJoinersWebPart extends BaseClientSideWebPart<INewJoinersWebPartProps> {

    private _siteListService!: SiteListService;

    protected async onInit(): Promise<void> {
        await super.onInit();
        this._siteListService = new SiteListService(this.context);
        ThemeService.initialize(this.context);

        if (!this.properties.siteUrl) {
            this.properties.siteUrl = this.context.pageContext.web.absoluteUrl;
        }

        if (!this.properties.manualJoiners) {
            this.properties.manualJoiners = [];
        }
    }

    public render(): void {
        const element: React.ReactElement<INewJoinersProps> = React.createElement(
            NewJoiners,
            {
                ...this.properties,
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

    private _addJoiner(): void {
        this.properties.manualJoiners.push({ user: null, introText: '' });
        this.context.propertyPane.refresh();
    }

    private _removeJoiner(index: number): void {
        this.properties.manualJoiners.splice(index, 1);
        this.context.propertyPane.refresh();
        this.render();
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        const manualJoinerFields: any[] = [];

        if (this.properties.source === 'graph') {
            this.properties.manualJoiners.forEach((joiner, index) => {
                manualJoinerFields.push(
                    PropertyFieldPeoplePicker(`manualJoiners[${index}].user`, {
                        label: `Person ${index + 1}`,
                        wpContext: this.context,
                        properties: this.properties,
                        onPropertyChange: (path, old, val) => {
                            this.properties.manualJoiners[index].user = (val && val.length > 0) ? val[0] : null;
                            this.render();
                        },
                        key: `peoplePicker-${index}`,
                        itemLimit: 1
                    }),
                    PropertyPaneTextField(`manualJoiners[${index}].introText`, {
                        label: 'Welcome Greeting',
                        value: joiner.introText,
                        onGetErrorMessage: (val) => {
                            this.properties.manualJoiners[index].introText = val;
                            this.render();
                            return '';
                        }
                    }),
                    PropertyPaneButton(`remove-${index}`, {
                        text: 'Remove',
                        buttonType: PropertyPaneButtonType.Command,
                        onClick: this._removeJoiner.bind(this, index)
                    }),
                    PropertyPaneHorizontalRule()
                );
            });

            manualJoinerFields.push(
                PropertyPaneButton('addJoiner', {
                    text: 'Add New Joiner',
                    buttonType: PropertyPaneButtonType.Primary,
                    onClick: this._addJoiner.bind(this)
                })
            );
        }

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
                                PropertyPaneChoiceGroup('source', {
                                    label: 'Select Data Source',
                                    options: [
                                        { key: 'spList', text: 'SharePoint List', iconProps: { officeFabricIconFontName: 'List' } },
                                        { key: 'graph', text: 'Microsoft Graph (Manual)', iconProps: { officeFabricIconFontName: 'People' } }
                                    ]
                                }),
                                ...(this.properties.source === 'spList' ? [
                                    PropertyFieldSitePicker('siteUrl', {
                                        label: 'Select Site',
                                        siteListService: this._siteListService,
                                        onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                        properties: this.properties,
                                        wpContext: this.context,
                                        key: 'njSiteUrlPicker'
                                    }),
                                    PropertyFieldListPicker('listId', {
                                        label: 'Choose Employee List',
                                        siteUrl: this.properties.siteUrl,
                                        siteListService: this._siteListService,
                                        onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                        properties: this.properties,
                                        wpContext: this.context,
                                        key: 'njListIdPicker',
                                        disabled: !this.properties.siteUrl
                                    })
                                ] : [])
                            ]
                        },
                        {
                            groupName: this.properties.source === 'graph' ? 'Manual New Joiners' : strings.ColumnMappingGroupName,
                            groupFields: this.properties.source === 'graph' ? manualJoinerFields : [
                                this._getColumnPicker('nameColumn', 'Name Column', 'Text', 'njName'),
                                this._getColumnPicker('photoColumn', 'Photo Column (Optional)', 'Thumbnail,URL,Image', 'njPhoto'),
                                this._getColumnPicker('jobTitleColumn', 'Job Title Column', 'Text', 'njJob'),
                                this._getColumnPicker('departmentColumn', 'Department Column', 'Text', 'njDept'),
                                this._getColumnPicker('emailColumn', 'Email/Person Column', 'User', 'njEmail'),
                                this._getColumnPicker('newJoinerColumn', 'New Joiner Flag (Yes/No)', 'Boolean', 'njFlag'),
                                this._getColumnPicker('newJoinerTextColumn', 'Welcome Intro Column', 'Text,Note', 'njText')
                            ]
                        },
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('webPartTitle', {
                                    label: 'Web Part Title'
                                }),
                                PropertyPaneDropdown('layout', {
                                    label: 'Layout View',
                                    options: [
                                        { key: 'list', text: 'List View' },
                                        { key: 'grid', text: 'Grid View (Cards)' },
                                        { key: 'strip', text: 'Horizontal Strip' }
                                    ]
                                }),
                                PropertyPaneChoiceGroup('layoutMode', {
                                    label: 'Layout Mode',
                                    options: [
                                        { key: 'standard', text: 'Standard (Wide)', iconProps: { officeFabricIconFontName: 'WebAppBuilderFragment' } },
                                        { key: 'compact', text: 'Compact (Side Column)', iconProps: { officeFabricIconFontName: 'WebAppBuilderSlot' } }
                                    ]
                                }),
                                PropertyPaneSlider('maxItems', {
                                    label: 'Maximum items to display',
                                    min: 1,
                                    max: 10,
                                    step: 1
                                }),
                                ...(this.properties.layout === 'strip' ? [
                                    PropertyPaneSlider('autoRotateInterval', {
                                        label: 'Auto-Rotate Interval (seconds)',
                                        min: 3,
                                        max: 15,
                                        step: 1
                                    })
                                ] : []),
                                PropertyPaneToggle('showBackgroundBar', {
                                    label: 'Show Accent Bar',
                                    onText: 'Show',
                                    offText: 'Hide'
                                }),
                                PropertyPaneChoiceGroup('titleBarStyle', {
                                    label: 'Accent Bar Style',
                                    options: [
                                        { key: 'solid', text: 'Solid Background' },
                                        { key: 'underline', text: 'Underline' }
                                    ]
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
