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
import { PeoplePicker } from '../controls/PeoplePicker/PeoplePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IPersonaProps } from '@fluentui/react/lib/Persona';

export interface IPropertyFieldPeoplePickerProps {
    label: string;
    wpContext: WebPartContext;
    selectedItems?: IPersonaProps[];
    onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
    properties: any;
    key: string;
    itemLimit?: number;
    disabled?: boolean;
}

export interface IPropertyFieldPeoplePickerPropsInternal extends IPropertyFieldPeoplePickerProps, IPropertyPaneCustomFieldProps { }

class PropertyFieldPeoplePickerBuilder implements IPropertyPaneField<IPropertyFieldPeoplePickerPropsInternal> {
    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    public targetProperty: string;
    public properties: IPropertyFieldPeoplePickerPropsInternal;

    public constructor(_targetProperty: string, _properties: IPropertyFieldPeoplePickerPropsInternal) {
        this.targetProperty = _targetProperty;
        this.properties = _properties;
        this.properties.onRender = this._render.bind(this);
        this.properties.onDispose = this._dispose.bind(this);
    }

    private _render(elem: HTMLElement, context?: any, changeCallback?: (targetProperty?: string, newValue?: any) => void): void {
        const element: React.ReactElement = React.createElement(PeoplePicker, {
            context: this.properties.wpContext,
            label: this.properties.label,
            selectedItems: this.properties.properties[this.targetProperty] || [],
            onChange: (items?: IPersonaProps[]) => {
                const oldValue = this.properties.properties[this.targetProperty];
                this.properties.properties[this.targetProperty] = items;
                this.properties.onPropertyChange(this.targetProperty, oldValue, items);
                if (changeCallback) {
                    changeCallback(this.targetProperty, items);
                }
            },
            itemLimit: this.properties.itemLimit,
            disabled: this.properties.disabled
        });
        ReactDOM.render(element, elem);
    }

    private _dispose(elem: HTMLElement): void {
        ReactDOM.unmountComponentAtNode(elem);
    }
}

export function PropertyFieldPeoplePicker(targetProperty: string, properties: IPropertyFieldPeoplePickerProps): IPropertyPaneField<IPropertyFieldPeoplePickerPropsInternal> {
    return new PropertyFieldPeoplePickerBuilder(targetProperty, {
        ...properties,
        onRender: (elem: HTMLElement, context?: any, changeCallback?: (targetProperty?: string, newValue?: any) => void) => {
            // Placeholder
        },
        onDispose: (elem: HTMLElement) => {
            // Placeholder
        }
    });
}
