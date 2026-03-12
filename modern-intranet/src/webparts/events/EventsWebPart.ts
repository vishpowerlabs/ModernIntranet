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
import * as strings from 'EventsWebPartStrings';
import { Events } from './components/Events';
import { IEventsProps } from './components/IEventsProps';
import { SiteListService } from '../../common/services/SiteListService';
import { ThemeService } from '../../common/services/ThemeService';
import { PropertyFieldSitePicker } from '../../common/propertyPaneControls/PropertyFieldSitePicker';
import { PropertyFieldListPicker } from '../../common/propertyPaneControls/PropertyFieldListPicker';
import { PropertyFieldColumnPicker } from '../../common/propertyPaneControls/PropertyFieldColumnPicker';

export interface IEventsWebPartProps {
  siteUrl: string;
  listId: string;
  titleColumn: string;
  dateColumn: string;
  activeColumn: string;
  imageColumn: string;
  linkColumn: string;
  locationColumn: string;
  descriptionColumn: string;
  pinnedColumn: string;
  maxItems: number;
  itemsPerRow: number;
  showViewAll: boolean;
  viewAllUrl: string;
  showTitle: boolean;
  title: string;
  showBackgroundBar: boolean;
  titleBarStyle: 'solid' | 'underline';
  layout: 'list' | 'grid' | 'compact';
  showEventImage: boolean;
  showPagination: boolean;
}

export default class EventsWebPart extends BaseClientSideWebPart<IEventsWebPartProps> {

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
    const element: React.ReactElement<IEventsProps> = React.createElement(
      Events,
      {
        siteUrl: this.properties.siteUrl,
        listId: this.properties.listId,
        titleColumn: this.properties.titleColumn,
        dateColumn: this.properties.dateColumn,
        imageColumn: this.properties.imageColumn,
        linkColumn: this.properties.linkColumn,
        locationColumn: this.properties.locationColumn,
        descriptionColumn: this.properties.descriptionColumn,
        activeColumn: this.properties.activeColumn,
        pinnedColumn: this.properties.pinnedColumn,
        maxItems: this.properties.maxItems,
        itemsPerRow: this.properties.itemsPerRow || 4,
        showViewAll: this.properties.showViewAll,
        viewAllUrl: this.properties.viewAllUrl,
        showTitle: this.properties.showTitle,
        title: this.properties.title,
        showBackgroundBar: this.properties.showBackgroundBar,
        titleBarStyle: this.properties.titleBarStyle || 'underline',
        layout: this.properties.layout || 'grid',
        showEventImage: this.properties.showEventImage !== false,
        showPagination: !!this.properties.showPagination,
        siteId: this.context.pageContext.site.id.toString(),
        webId: this.context.pageContext.web.id.toString(),
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
                PropertyFieldColumnPicker('dateColumn', {
                  label: strings.DateColumnFieldLabel,
                  siteUrl: this.properties.siteUrl,
                  listId: this.properties.listId,
                  typeFilter: 'DateTime',
                  siteListService: this._siteListService,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  wpContext: this.context,
                  key: 'dateColumnPicker',
                  disabled: !this.properties.listId
                }),
                PropertyFieldColumnPicker('activeColumn', {
                  label: strings.ActiveColumnFieldLabel,
                  siteUrl: this.properties.siteUrl,
                  listId: this.properties.listId,
                  typeFilter: 'Boolean,Choice,Bit',
                  siteListService: this._siteListService,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  wpContext: this.context,
                  key: 'activeColumnPicker',
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
                }),
                PropertyFieldColumnPicker('locationColumn', {
                  label: strings.LocationColumnFieldLabel,
                  siteUrl: this.properties.siteUrl,
                  listId: this.properties.listId,
                  typeFilter: 'Text',
                  siteListService: this._siteListService,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  wpContext: this.context,
                  key: 'locationColumnPicker',
                  disabled: !this.properties.listId
                }),
                PropertyFieldColumnPicker('descriptionColumn', {
                  label: strings.DescriptionColumnFieldLabel,
                  siteUrl: this.properties.siteUrl,
                  listId: this.properties.listId,
                  typeFilter: 'Note,Text',
                  siteListService: this._siteListService,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  wpContext: this.context,
                  key: 'descriptionColumnPicker',
                  disabled: !this.properties.listId
                })
              ]
            },
            {
              groupName: strings.DisplaySettingsGroupName,
              groupFields: [
                PropertyPaneSlider('maxItems', {
                  label: strings.MaxItemsFieldLabel,
                  min: 2,
                  max: 12,
                  step: 1
                }),
                PropertyPaneSlider('itemsPerRow', {
                  label: strings.ItemsPerRowFieldLabel,
                  min: 2,
                  max: 4,
                  step: 1
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
                ] : []),
                PropertyPaneToggle('showViewAll', {
                  label: strings.ShowViewAllFieldLabel
                }),
                PropertyPaneTextField('viewAllUrl', {
                  label: strings.ViewAllUrlFieldLabel,
                  disabled: !this.properties.showViewAll
                }),
                PropertyPaneChoiceGroup('layout', {
                  label: 'Layout',
                  options: [
                    { key: 'grid', text: 'Grid View', iconProps: { officeFabricIconFontName: 'GridViewSmall' } },
                    { key: 'list', text: 'List View', iconProps: { officeFabricIconFontName: 'List' } },
                    { key: 'compact', text: 'Compact View', iconProps: { officeFabricIconFontName: 'TwelvePointStar' } }
                  ]
                }),
                PropertyPaneToggle('showPagination', {
                  label: strings.ShowPaginationFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
