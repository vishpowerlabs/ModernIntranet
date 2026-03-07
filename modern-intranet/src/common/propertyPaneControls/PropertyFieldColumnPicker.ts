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
import { ColumnPicker } from '../controls/ColumnPicker/ColumnPicker';
import { ISiteListService } from '../services/ISiteListService';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IPropertyFieldColumnPickerProps {
    label: string;
    siteUrl: string;
    listId: string;
    typeFilter?: string;
    siteListService: ISiteListService;
    onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
    properties: any;
    wpContext: WebPartContext;
    key: string;
    disabled?: boolean;
}

export interface IPropertyFieldColumnPickerPropsInternal extends IPropertyFieldColumnPickerProps, IPropertyPaneCustomFieldProps { }

class PropertyFieldColumnPickerBuilder implements IPropertyPaneField<IPropertyFieldColumnPickerPropsInternal> {
    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    public targetProperty: string;
    public properties: IPropertyFieldColumnPickerPropsInternal;
    private elem!: HTMLElement;

    public constructor(_targetProperty: string, _properties: IPropertyFieldColumnPickerPropsInternal) {
        this.targetProperty = _targetProperty;
        this.properties = _properties;
        this.properties.onRender = this._render.bind(this);
        this.properties.onDispose = this._dispose.bind(this);
    }

    private _render(elem: HTMLElement, context?: WebPartContext, changeCallback?: (targetProperty?: string, newValue?: any) => void): void {
        this.elem = elem;
        const element: React.ReactElement = React.createElement(ColumnPicker, {
            context: this.properties.wpContext,
            label: this.properties.label,
            siteUrl: this.properties.siteUrl,
            listId: this.properties.listId,
            typeFilter: this.properties.typeFilter,
            selectedColumn: this.properties.properties[this.targetProperty],
            disabled: this.properties.disabled,
            onColumnSelected: (columnInternalName: string) => {
                const oldValue = this.properties.properties[this.targetProperty];
                this.properties.properties[this.targetProperty] = columnInternalName;
                this.properties.onPropertyChange(this.targetProperty, oldValue, columnInternalName);
                if (changeCallback) {
                    changeCallback(this.targetProperty, columnInternalName);
                }
            }
        });
        ReactDOM.render(element, elem);
    }

    private _dispose(elem: HTMLElement): void {
        ReactDOM.unmountComponentAtNode(elem);
    }
}

export function PropertyFieldColumnPicker(targetProperty: string, properties: IPropertyFieldColumnPickerProps): IPropertyPaneField<IPropertyFieldColumnPickerPropsInternal> {
    return new PropertyFieldColumnPickerBuilder(targetProperty, {
        ...properties,
        onRender: (elem: HTMLElement, context?: WebPartContext, changeCallback?: (targetProperty?: string, newValue?: any) => void) => {
            // Placeholder
        },
        onDispose: (elem: HTMLElement) => {
            // Placeholder
        }
    });
}
