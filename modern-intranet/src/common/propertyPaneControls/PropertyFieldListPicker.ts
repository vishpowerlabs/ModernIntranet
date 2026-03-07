import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {
    IPropertyPaneField,
    PropertyPaneFieldType,
    IPropertyPaneCustomFieldProps
} from '@microsoft/sp-property-pane';
import { ListPicker } from '../controls/ListPicker/ListPicker';
import { ISiteListService } from '../services/ISiteListService';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IPropertyFieldListPickerProps {
    label: string;
    siteUrl: string;
    siteListService: ISiteListService;
    onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
    properties: any;
    wpContext: WebPartContext;
    key: string;
    disabled?: boolean;
}

export interface IPropertyFieldListPickerPropsInternal extends IPropertyFieldListPickerProps, IPropertyPaneCustomFieldProps { }

class PropertyFieldListPickerBuilder implements IPropertyPaneField<IPropertyFieldListPickerPropsInternal> {
    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    public targetProperty: string;
    public properties: IPropertyFieldListPickerPropsInternal;
    private elem!: HTMLElement;

    public constructor(_targetProperty: string, _properties: IPropertyFieldListPickerPropsInternal) {
        this.targetProperty = _targetProperty;
        this.properties = _properties;
        this.properties.onRender = this._render.bind(this);
        this.properties.onDispose = this._dispose.bind(this);
    }

    private _render(elem: HTMLElement, context?: WebPartContext, changeCallback?: (targetProperty?: string, newValue?: string) => void): void {
        this.elem = elem;
        const element: React.ReactElement = React.createElement(ListPicker, {
            context: this.properties.wpContext,
            label: this.properties.label,
            siteUrl: this.properties.siteUrl,
            selectedListId: this.properties.properties[this.targetProperty],
            onListSelected: (listId: string) => {
                const oldValue = this.properties.properties[this.targetProperty];
                this.properties.properties[this.targetProperty] = listId;
                this.properties.onPropertyChange(this.targetProperty, oldValue, listId);
                if (changeCallback) {
                    changeCallback(this.targetProperty, listId);
                }
            }
        });
        ReactDOM.render(element, elem);
    }

    private _dispose(elem: HTMLElement): void {
        ReactDOM.unmountComponentAtNode(elem);
    }
}

export function PropertyFieldListPicker(targetProperty: string, properties: IPropertyFieldListPickerProps): IPropertyPaneField<IPropertyFieldListPickerPropsInternal> {
    return new PropertyFieldListPickerBuilder(targetProperty, {
        ...properties,
        onRender: (elem: HTMLElement, context?: WebPartContext, changeCallback?: (targetProperty?: string, newValue?: string) => void) => {
            // Placeholder
        },
        onDispose: (elem: HTMLElement) => {
            // Placeholder
        }
    });
}
