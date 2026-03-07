/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {
    IPropertyPaneField,
    PropertyPaneFieldType,
    IPropertyPaneCustomFieldProps
} from '@microsoft/sp-property-pane';
import { SitePicker } from '../controls/SitePicker/SitePicker';
import { ISiteListService } from '../services/ISiteListService';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IPropertyFieldSitePickerProps {
    label: string;
    siteListService: ISiteListService;
    initialSiteUrl?: string;
    onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
    properties: any;
    wpContext: WebPartContext;
    key: string;
}

export interface IPropertyFieldSitePickerPropsInternal extends IPropertyFieldSitePickerProps, IPropertyPaneCustomFieldProps { }

class PropertyFieldSitePickerBuilder implements IPropertyPaneField<IPropertyFieldSitePickerPropsInternal> {
    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    public targetProperty: string;
    public properties: IPropertyFieldSitePickerPropsInternal;
    private elem!: HTMLElement;

    public constructor(_targetProperty: string, _properties: IPropertyFieldSitePickerPropsInternal) {
        this.targetProperty = _targetProperty;
        this.properties = _properties;
        this.properties.onRender = this._render.bind(this);
        this.properties.onDispose = this._dispose.bind(this);
    }

    private _render(elem: HTMLElement, context?: WebPartContext, changeCallback?: (targetProperty?: string, newValue?: any) => void): void {
        this.elem = elem;
        const element: React.ReactElement = React.createElement(SitePicker, {
            label: this.properties.label,
            context: this.properties.wpContext, // Passing the context down
            selectedSiteUrl: this.properties.properties[this.targetProperty],
            onSiteSelected: (siteUrl: string) => {
                const oldValue = this.properties.properties[this.targetProperty];
                this.properties.properties[this.targetProperty] = siteUrl;
                this.properties.onPropertyChange(this.targetProperty, oldValue, siteUrl);
                if (changeCallback) {
                    changeCallback(this.targetProperty, siteUrl);
                }
            }
        });
        ReactDOM.render(element, elem);
    }

    private _dispose(elem: HTMLElement): void {
        ReactDOM.unmountComponentAtNode(elem);
    }
}

export function PropertyFieldSitePicker(targetProperty: string, properties: IPropertyFieldSitePickerProps): IPropertyPaneField<IPropertyFieldSitePickerPropsInternal> {
    return new PropertyFieldSitePickerBuilder(targetProperty, {
        ...properties,
        onRender: (elem: HTMLElement, context?: WebPartContext, changeCallback?: (targetProperty?: string, newValue?: any) => void) => {
            // Placeholder, overwritten in constructor
        },
        onDispose: (elem: HTMLElement) => {
            // Placeholder, overwritten in constructor
        }
    });
}
