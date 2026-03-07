import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PropertyFieldSitePicker, IPropertyFieldSite } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';

import * as strings from 'DocumentListingV2WebPartStrings';
import { DocumentListingV2 } from './components/DocumentListingV2';
import { IDocumentListingV2Props } from './components/IDocumentListingV2Props';
import {
  ISPListsResponse,
  ISPFieldsResponse,
  ISPList,
  ISPField
} from './models/ISPData';

export interface IDocumentListingV2WebPartProps {
  description: string;
  sourceLibraryId: string;
  requestListId: string;
  categoryField: string;
  subCategoryField: string;
  descriptionField: string;
  pageSize: number;
  requestEmailField: string;
  requestFileIdField: string;
  requestRequestIdField: string;
  requestDateField: string;
  requestReminderField: string;
  alreadyRequestedMessage: string;
  webPartTitle: string;
  webPartTitleFontSize: string;
  webPartDescription: string;
  webPartDescriptionFontSize: string;
  reminderSentMessage: string;
  headerOpacity: number;
  pinnedField: string;
  showRequestAccess: boolean;
  sites: IPropertyFieldSite[];
  requestSites: IPropertyFieldSite[];
}

export default class DocumentListingV2WebPart extends BaseClientSideWebPart<IDocumentListingV2WebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _themePrimary: string = '';
  private _bodyBackground: string = '#ffffff';
  private _listsDropdownOptions: IPropertyPaneDropdownOption[] = [];
  private _requestListOptions: IPropertyPaneDropdownOption[] = [];
  private _fieldsDropdownOptions: IPropertyPaneDropdownOption[] = [];
  private _requestFieldsDropdownOptions: IPropertyPaneDropdownOption[] = [];

  public render(): void {
    const siteUrl = this.properties.sites && this.properties.sites.length > 0 ? this.properties.sites[0].url : this.context.pageContext.web.absoluteUrl;
    const requestSiteUrl = this.properties.requestSites && this.properties.requestSites.length > 0 ? this.properties.requestSites[0].url : this.context.pageContext.web.absoluteUrl;

    const element: React.ReactElement<IDocumentListingV2Props> = React.createElement(
      DocumentListingV2,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        sourceLibraryId: this.properties.sourceLibraryId,
        requestListId: this.properties.requestListId,
        categoryField: this.properties.categoryField || 'Category',
        subCategoryField: this.properties.subCategoryField || 'SubCategory',
        descriptionField: this.properties.descriptionField || 'Description',
        pageSize: this.properties.pageSize || 10,
        requestEmailField: this.properties.requestEmailField || 'Email',
        requestFileIdField: this.properties.requestFileIdField || 'FileID',
        requestRequestIdField: this.properties.requestRequestIdField || 'RequestID',
        requestDateField: this.properties.requestDateField || 'RequestDate',
        requestReminderField: this.properties.requestReminderField || 'Reminder',
        alreadyRequestedMessage: this.properties.alreadyRequestedMessage || 'You have already requested access. Please check your email.',
        webPartTitle: this.properties.webPartTitle,
        webPartTitleFontSize: this.properties.webPartTitleFontSize || '24px',
        webPartDescription: this.properties.webPartDescription,
        webPartDescriptionFontSize: this.properties.webPartDescriptionFontSize || '14px',
        reminderSentMessage: this.properties.reminderSentMessage || 'Reminder sent successfully!',
        headerOpacity: this.properties.headerOpacity !== undefined ? this.properties.headerOpacity : 1,
        headerTextColor: this._getContrastTextColor(
          this._themePrimary || '#0078d4',
          this.properties.headerOpacity !== undefined ? this.properties.headerOpacity : 1,
          this._bodyBackground
        ),
        pinnedField: this.properties.pinnedField,
        showRequestAccess: this.properties.showRequestAccess !== undefined ? this.properties.showRequestAccess : true,
        siteUrl: siteUrl,
        requestSiteUrl: requestSiteUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
      this._fetchLists();
      this._fetchLists();
      // If we already have a library selected, fetch fields
      const siteUrl = this.properties.sites && this.properties.sites.length > 0 ? this.properties.sites[0].url : this.context.pageContext.web.absoluteUrl;
      const requestSiteUrl = this.properties.requestSites && this.properties.requestSites.length > 0 ? this.properties.requestSites[0].url : this.context.pageContext.web.absoluteUrl;

      if (this.properties.sourceLibraryId) {
        this._fetchFields(this.properties.sourceLibraryId, siteUrl);
      }
      if (this.properties.requestListId) {
        this._fetchRequestFields(this.properties.requestListId, requestSiteUrl);
      }
      this._fetchRequestLists(requestSiteUrl);
    });
  }

  private _fetchLists(siteUrl?: string): void {
    const webUrl = siteUrl || this.context.pageContext.web.absoluteUrl;
    this.context.spHttpClient.get(
      `${webUrl}/_api/web/lists?$select=Id,Title&$filter=Hidden eq false`,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: ISPListsResponse) => {
        this._listsDropdownOptions = data.value.map((list: ISPList) => {
          return {
            key: list.Id,
            text: list.Title
          };
        });
        this.context.propertyPane.refresh();
      })
      .catch((error) => {
        console.error('Error fetching lists', error);
      });
  }

  private _fetchRequestLists(siteUrl?: string): void {
    const webUrl = siteUrl || this.context.pageContext.web.absoluteUrl;
    this.context.spHttpClient.get(
      `${webUrl}/_api/web/lists?$select=Id,Title&$filter=Hidden eq false`,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: ISPListsResponse) => {
        this._requestListOptions = data.value.map((list: ISPList) => {
          return {
            key: list.Id,
            text: list.Title
          };
        });
        this.context.propertyPane.refresh();
      })
      .catch((error) => {
        console.error('Error fetching request lists', error);
      });
  }

  private _fetchFields(listId: string, siteUrl?: string): void {
    if (!listId) return;

    const webUrl = siteUrl || this.context.pageContext.web.absoluteUrl;

    this.context.spHttpClient.get(
      `${webUrl}/_api/web/lists(guid'${listId}')/fields?$filter=Hidden eq false and ReadOnlyField eq false`,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: ISPFieldsResponse) => {
        this._fieldsDropdownOptions = data.value.map((field: ISPField) => {
          return {
            key: field.InternalName,
            text: `${field.Title} (${field.InternalName})`
          };
        });
        // Sort alphabetically
        this._fieldsDropdownOptions.sort((a, b) => a.text.localeCompare(b.text));

        this.context.propertyPane.refresh();
      })
      .catch((error) => {
        console.error('Error fetching fields', error);
      });
  }

  private _fetchRequestFields(listId: string, siteUrl?: string): void {
    if (!listId) return;

    const webUrl = siteUrl || this.context.pageContext.web.absoluteUrl;

    this.context.spHttpClient.get(
      `${webUrl}/_api/web/lists(guid'${listId}')/fields?$filter=Hidden eq false and ReadOnlyField eq false`,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: ISPFieldsResponse) => {
        this._requestFieldsDropdownOptions = data.value.map((field: ISPField) => {
          return {
            key: field.InternalName,
            text: `${field.Title} (${field.InternalName})`
          };
        });
        this._requestFieldsDropdownOptions.sort((a, b) => a.text.localeCompare(b.text));
        this.context.propertyPane.refresh();
      })
      .catch((error) => {
        console.error('Error fetching request fields', error);
      });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    const siteUrl = this.properties.sites && this.properties.sites.length > 0 ? this.properties.sites[0].url : this.context.pageContext.web.absoluteUrl;
    const requestSiteUrl = this.properties.requestSites && this.properties.requestSites.length > 0 ? this.properties.requestSites[0].url : this.context.pageContext.web.absoluteUrl;

    if (propertyPath === 'sites') {
      this._listsDropdownOptions = [];
      this._fieldsDropdownOptions = [];

      this.properties.sourceLibraryId = '';
      this.properties.categoryField = '';
      this.properties.subCategoryField = '';
      this.properties.descriptionField = '';
      this.properties.pinnedField = '';

      this._fetchLists(siteUrl);
    }

    if (propertyPath === 'requestSites') {
      this._requestFieldsDropdownOptions = [];

      this.properties.requestListId = '';
      this.properties.requestEmailField = '';
      this.properties.requestFileIdField = '';
      this.properties.requestRequestIdField = '';
      this.properties.requestDateField = '';
      this.properties.requestReminderField = '';

      // Refetch lists for requests
      this._fetchRequestLists(requestSiteUrl);
    }

    if (propertyPath === 'sourceLibraryId' && newValue) {
      this._fieldsDropdownOptions = [];
      this._fetchFields(newValue, siteUrl);
    }

    if (propertyPath === 'requestListId' && newValue) {
      this._requestFieldsDropdownOptions = [];
      this._fetchRequestFields(newValue, requestSiteUrl);
    }
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    if (currentTheme) {
      if (currentTheme.palette) {
        this._themePrimary = currentTheme.palette.themePrimary || '';
        this.domElement.style.setProperty('--themePrimary', this._themePrimary);
      }

      const { semanticColors } = currentTheme;
      if (semanticColors) {
        this._bodyBackground = semanticColors.bodyBackground || '#ffffff';
        this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || '');
        this.domElement.style.setProperty('--link', semanticColors.link || '');
        this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || '');
      }
    }
  }

  private _getContrastTextColor(color: string, opacity: number, baseColor: string): string {
    // Helper to parse hex to rgb
    const hexToRgb = (hex: string): { r: number, g: number, b: number } | null => {
      const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
      return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
      } : null;
    };

    const fg = hexToRgb(color) || { r: 0, g: 120, b: 212 }; // Default blue
    const bg = hexToRgb(baseColor) || { r: 255, g: 255, b: 255 }; // Default white

    // Mix colors based on opacity
    const r = Math.round(fg.r * opacity + bg.r * (1 - opacity));
    const g = Math.round(fg.g * opacity + bg.g * (1 - opacity));
    const b = Math.round(fg.b * opacity + bg.b * (1 - opacity));

    // Calculate YIQ
    const yiq = ((r * 299) + (g * 587) + (b * 114)) / 1000;
    return (yiq >= 128) ? 'var(--bodyText)' : '#ffffff';
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
            description: "Configure Document Listing"
          },
          groups: [
            {
              groupName: "Display Settings",
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
                  ],
                  selectedKey: '24px'
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
                  ],
                  selectedKey: '14px'
                }),
                PropertyPaneSlider('headerOpacity', {
                  label: 'Header Opacity',
                  min: 0,
                  max: 1,
                  step: 0.1,
                  showValue: true
                })
              ]
            },
            {
              groupName: "Document Library Settings",
              groupFields: [
                PropertyFieldSitePicker('sites', {
                  label: 'Select Site',
                  initialSites: this.properties.sites,
                  context: this.context as any,
                  deferredValidationTime: 500,
                  multiSelect: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'sitesField'
                }),
                PropertyPaneDropdown('sourceLibraryId', {
                  label: 'Source Document Library',
                  options: this._listsDropdownOptions,
                  disabled: this._listsDropdownOptions.length === 0
                }),
                PropertyPaneDropdown('categoryField', {
                  label: 'Category Field',
                  options: this._fieldsDropdownOptions,
                  disabled: this._fieldsDropdownOptions.length === 0
                }),
                PropertyPaneDropdown('subCategoryField', {
                  label: 'Sub-Category Field',
                  options: this._fieldsDropdownOptions,
                  disabled: this._fieldsDropdownOptions.length === 0
                }),
                PropertyPaneDropdown('descriptionField', {
                  label: 'Description Field',
                  options: this._fieldsDropdownOptions,
                  disabled: this._fieldsDropdownOptions.length === 0
                }),
                PropertyPaneTextField('pageSize', {
                  label: 'Items per page',
                  description: 'Number of items to show per page'
                }),
                PropertyPaneDropdown('pinnedField', {
                  label: 'Pinned Column',
                  options: this._fieldsDropdownOptions,
                  disabled: this._fieldsDropdownOptions.length === 0
                })
              ]
            },
            {
              groupName: "Request Access Settings",
              groupFields: [
                PropertyPaneToggle('showRequestAccess', {
                  label: 'Show Request Access Column',
                  onText: 'Show',
                  offText: 'Hide'
                }),
                PropertyFieldSitePicker('requestSites', {
                  label: 'Select Request Site',
                  initialSites: this.properties.requestSites,
                  context: this.context as any,
                  deferredValidationTime: 500,
                  multiSelect: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'requestSitesField'
                }),
                PropertyPaneDropdown('requestListId', {
                  label: 'Requests List',
                  options: this._requestListOptions,
                  disabled: this._requestListOptions.length === 0
                }),
                PropertyPaneDropdown('requestEmailField', {
                  label: 'Email Field (in Request List)',
                  options: this._requestFieldsDropdownOptions,
                  disabled: this._requestFieldsDropdownOptions.length === 0
                }),
                PropertyPaneDropdown('requestFileIdField', {
                  label: 'File ID Field',
                  options: this._requestFieldsDropdownOptions,
                  disabled: this._requestFieldsDropdownOptions.length === 0
                }),
                PropertyPaneDropdown('requestRequestIdField', {
                  label: 'Request ID Field',
                  options: this._requestFieldsDropdownOptions,
                  disabled: this._requestFieldsDropdownOptions.length === 0
                }),
                PropertyPaneDropdown('requestDateField', {
                  label: 'Request Date Field',
                  options: this._requestFieldsDropdownOptions,
                  disabled: this._requestFieldsDropdownOptions.length === 0
                }),
                PropertyPaneDropdown('requestReminderField', {
                  label: 'Reminder Field (e.g., SendAgain)',
                  options: this._requestFieldsDropdownOptions,
                  disabled: this._requestFieldsDropdownOptions.length === 0
                }),
                PropertyPaneTextField('alreadyRequestedMessage', {
                  label: 'Already Requested Message',
                  description: 'Message to show if user has already requested access',
                  multiline: true
                }),
                PropertyPaneTextField('reminderSentMessage', {
                  label: 'Reminder Sent Message',
                  description: 'Toast message when reminder is sent'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
