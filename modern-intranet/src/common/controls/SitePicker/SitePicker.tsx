/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { useSites } from '../../hooks/useSites';
import styles from './SitePicker.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISitePickerProps {
    context: WebPartContext;
    selectedSiteUrl: string;
    onSiteSelected: (siteUrl: string) => void;
    label?: string;
    required?: boolean;
}

export const SitePicker: React.FC<ISitePickerProps> = ({
    context,
    selectedSiteUrl,
    onSiteSelected,
    label = 'Select a Site',
    required = false
}) => {
    const sites = useSites(context);

    const options: IDropdownOption[] = sites.map(site => ({
        key: site.url,
        text: site.title,
        data: site
    }));

    React.useEffect(() => {
        // Default to current site if none selected and sites are loaded
        if (!selectedSiteUrl && sites.length > 0 && context?.pageContext?.web?.absoluteUrl) {
            const currentUrl = context.pageContext.web.absoluteUrl;
            const currentSiteExists = sites.some(s => s.url === currentUrl);
            if (currentSiteExists) {
                onSiteSelected(currentUrl);
            }
        }
    }, [sites, selectedSiteUrl, context, onSiteSelected]);

    const onChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
        if (option) {
            onSiteSelected(option.key as string);
        }
    };

    return (
        <div className={styles.sitePicker}>
            <Dropdown
                label={label}
                selectedKey={selectedSiteUrl}
                onChange={onChange}
                placeholder="Select a site..."
                options={options}
                required={required}
            />
        </div>
    );
};
