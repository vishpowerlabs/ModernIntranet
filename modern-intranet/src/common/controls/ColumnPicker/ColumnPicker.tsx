/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { useColumns } from '../../hooks/useColumns';
import styles from './ColumnPicker.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IColumnPickerProps {
    context: WebPartContext;
    siteUrl: string;
    listId: string;
    selectedColumn: string;
    onColumnSelected: (columnInternalName: string) => void;
    typeFilter?: string;
    label?: string;
    required?: boolean;
    disabled?: boolean;
}

export const ColumnPicker: React.FC<IColumnPickerProps> = ({
    context,
    siteUrl,
    listId,
    selectedColumn,
    onColumnSelected,
    typeFilter,
    label = 'Select a Column',
    required = false,
    disabled = false
}) => {
    const columns = useColumns(context, siteUrl, listId, typeFilter);

    React.useEffect(() => {
        if (selectedColumn && columns.length > 0) {
            const exists = columns.some(c => c.internalName === selectedColumn);
            if (!exists) {
                onColumnSelected('');
            }
        }
    }, [columns, selectedColumn, onColumnSelected]);

    const options: IDropdownOption[] = columns.map(col => ({
        key: col.internalName,
        text: `${col.title} (${col.internalName})`,
        data: col
    }));

    const onChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
        if (option) {
            onColumnSelected(option.key as string);
        }
    };

    return (
        <div className={styles.columnPicker}>
            <Dropdown
                label={label}
                selectedKey={selectedColumn}
                onChange={onChange}
                placeholder={listId ? "Select a column..." : "Select a list first..."}
                options={options}
                disabled={disabled || !listId}
                required={required}
            />
        </div>
    );
};
