/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { useLists } from '../../hooks/useLists';
import styles from './ListPicker.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IListPickerProps {
    context: WebPartContext;
    siteUrl: string;
    selectedListId: string;
    onListSelected: (listId: string) => void;
    label?: string;
    required?: boolean;
}

export const ListPicker: React.FC<IListPickerProps> = ({
    context,
    siteUrl,
    selectedListId,
    onListSelected,
    label = 'Select a List',
    required = false
}) => {
    const lists = useLists(context, siteUrl);

    // Clear selection if the current siteUrl changes and the lists update, and the selected list is not in the new options.
    // Although simpler: Clear when siteUrl changes.
    React.useEffect(() => {
        if (selectedListId && lists.length > 0) {
            const exists = lists.some(l => l.id === selectedListId);
            if (!exists) {
                onListSelected('');
            }
        }
    }, [lists, selectedListId, onListSelected]);

    const options: IDropdownOption[] = lists.map(list => ({
        key: list.id,
        text: list.title,
        data: list
    }));

    const onChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
        if (option) {
            onListSelected(option.key as string);
        }
    };

    return (
        <div className={styles.listPicker}>
            <Dropdown
                label={label}
                selectedKey={selectedListId}
                onChange={onChange}
                placeholder={siteUrl ? "Select a list..." : "Select a site first..."}
                options={options}
                disabled={!siteUrl || lists.length === 0}
                required={required}
            />
        </div>
    );
};
