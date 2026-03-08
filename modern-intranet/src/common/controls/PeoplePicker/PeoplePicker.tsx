/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import {
    NormalPeoplePicker,
    IBasePickerSuggestionsProps
} from '@fluentui/react/lib/Pickers';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import styles from './PeoplePicker.module.scss';

export interface IPeoplePickerProps {
    context: WebPartContext;
    label?: string;
    selectedItems?: IPersonaProps[];
    onChange?: (items?: IPersonaProps[]) => void;
    itemLimit?: number;
    disabled?: boolean;
}

const suggestionProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: 'Suggested People',
    mostRecentlyUsedHeaderText: 'Suggested Contacts',
    noResultsFoundText: 'No results found',
    loadingText: 'Loading',
    showRemoveButtons: true,
};

export const PeoplePicker: React.FC<IPeoplePickerProps> = ({
    context,
    label,
    selectedItems,
    onChange,
    itemLimit = 10,
    disabled = false
}) => {

    const onFilterChanged = async (filterText: string, currentPersonas?: IPersonaProps[]): Promise<IPersonaProps[]> => {
        if (!filterText || filterText.length < 3) {
            return [];
        }

        try {
            const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient('3');
            const result = await client.api('/users')
                .filter(`startsWith(displayName,'${filterText}') or startsWith(givenName,'${filterText}') or startsWith(surname,'${filterText}') or startsWith(mail,'${filterText}')`)
                .select('id,displayName,jobTitle,department,mail,userPrincipalName')
                .top(10)
                .get();

            const personas: IPersonaProps[] = result.value.map((user: any) => ({
                id: user.id,
                text: user.displayName,
                secondaryText: user.mail || user.userPrincipalName,
                tertiaryText: user.jobTitle,
                optionalText: user.department,
                imageUrl: `/_layouts/15/userphoto.aspx?size=S&accountname=${user.mail || user.userPrincipalName}`
            }));

            // Filter out already selected personas
            return personas.filter(persona => !currentPersonas?.some(selected => selected.id === persona.id));
        } catch (error) {
            console.error('Error fetching users from Graph:', error);
            return [];
        }
    };

    return (
        <div className={styles.peoplePickerContainer}>
            {label && <label className={styles.label}>{label}</label>}
            <NormalPeoplePicker
                onResolveSuggestions={onFilterChanged}
                getTextFromItem={(persona: IPersonaProps) => persona.text || ''}
                pickerSuggestionsProps={suggestionProps}
                className={styles.picker}
                onChange={onChange}
                selectedItems={selectedItems}
                itemLimit={itemLimit}
                disabled={disabled}
            />
        </div>
    );
};
