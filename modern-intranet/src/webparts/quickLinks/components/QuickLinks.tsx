/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IQuickLinksProps } from './IQuickLinksProps';
import { LinkTile } from './LinkTile';
import { EmptyState } from '../../../common/components/EmptyState/EmptyState';
import styles from './QuickLinks.module.scss';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

interface ISharePointLinkValue {
    Url?: string;
    Description?: string;
}

interface ISharePointQuickLinkRow {
    Id: number;
    Title: string;
    [key: string]: string | number | boolean | undefined | null | ISharePointLinkValue;
}

export interface IQuickLinkItem {
    id: number;
    title: string;
    url: string;
    icon?: string;
    pinned: boolean;
}

export const QuickLinks: React.FC<IQuickLinksProps> = (props) => {
    const [items, setItems] = React.useState<IQuickLinkItem[]>([]);
    const [loading, setLoading] = React.useState<boolean>(false);
    const [error, setError] = React.useState<string | undefined>();

    const loadData = async (): Promise<void> => {
        if (!props.siteUrl || !props.listId) {
            setItems([]);
            return;
        }

        if (!props.titleColumn || !props.linkColumn) {
            setError('Please map Title and Link columns in the property pane.');
            return;
        }

        setLoading(true);
        setError(undefined);

        try {
            const selectFields = ['Id', props.titleColumn, props.linkColumn];
            if (props.iconColumn) {
                selectFields.push(props.iconColumn);
            }
            if (props.pinnedColumn) {
                selectFields.push(props.pinnedColumn);
            }

            const apiUrl = `${props.siteUrl.replace(/\/$/, '')}/_api/web/lists(guid'${props.listId}')/items?$select=${selectFields.join(',')}&$orderby=Title asc`;

            const response: SPHttpClientResponse = await props.context.spHttpClient.get(
                apiUrl,
                SPHttpClient.configurations.v1
            );

            if (!response.ok) {
                throw new Error(`Error fetching list items: ${response.statusText}`);
            }

            const data = await response.json();

            const mappedItems: IQuickLinkItem[] = (data.value || []).map((item: ISharePointQuickLinkRow) => {
                let urlValue = '';
                const rawLink = item[props.linkColumn];

                if (rawLink && typeof rawLink === 'object' && 'Url' in rawLink) {
                    urlValue = rawLink.Url || '';
                } else if (typeof rawLink === 'string') {
                    urlValue = rawLink;
                }

                return {
                    id: item.Id,
                    title: item[props.titleColumn] || '',
                    url: urlValue,
                    icon: props.iconColumn ? item[props.iconColumn] : undefined,
                    pinned: props.pinnedColumn ? !!item[props.pinnedColumn] : false
                };
            }).sort((a: IQuickLinkItem, b: IQuickLinkItem) => {
                if (a.pinned !== b.pinned) {
                    return a.pinned ? -1 : 1;
                }
                return String(a.title).localeCompare(String(b.title));
            });

            setItems(mappedItems);
        } catch (err) {
            const errorMessage = (err as Error).message || 'An error occurred loading links.';
            setError(errorMessage);
        } finally {
            setLoading(false);
        }
    };

    React.useEffect(() => {
        loadData().catch(err => console.error('Error in QuickLinks useEffect:', err));
    }, [props.siteUrl, props.listId, props.titleColumn, props.linkColumn, props.iconColumn, props.pinnedColumn]);

    const isConfigured = props.siteUrl && props.listId && props.titleColumn && props.linkColumn;

    if (!isConfigured) {
        return (
            <div className={styles.quickLinks}>
                <EmptyState
                    icon="Link"
                    title="Quick Links - Configuration Required"
                    message="Please complete the web part configuration to display tiles."
                    description="You need to specify the Site URL, List ID, and map the required columns (Title and Link) in the property pane."
                />
            </div>
        );
    }

    if (error) {
        return (
            <div className={styles.quickLinks}>
                <EmptyState icon="Error" message={error} />
            </div>
        );
    }

    const getHeaderClass = (): string => {
        if (!props.showBackgroundBar) return '';
        return props.titleBarStyle === 'solid' ? styles.solidBackground : styles.underlineBackground;
    };

    if (loading) {
        return (
            <div className={styles.quickLinks}>
                <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '200px' }}>
                    <Spinner size={SpinnerSize.large} label="Loading links..." />
                </div>
            </div>
        );
    }

    if (items.length === 0) {
        return (
            <div className={styles.quickLinks}>
                <EmptyState
                    icon="SearchIssue"
                    title="No Quick Links Found"
                    message="There are no links to display from the selected list."
                    description="Add items to your SharePoint list or check your filter settings if applicable."
                />
            </div>
        );
    }

    let columnsClass = styles.cols6;
    if (props.columnsPerRow === 2) {
        columnsClass = styles.cols2;
    } else if (props.columnsPerRow === 3) {
        columnsClass = styles.cols3;
    } else if (props.columnsPerRow === 4) {
        columnsClass = styles.cols4;
    }

    return (
        <div className={styles.quickLinks}>
            {props.showTitle && props.title && (
                <div className={`${styles.webpartHeader} ${getHeaderClass()}`}>
                    <div className={styles.titleContainer}>
                        <h2>{props.title}</h2>
                    </div>
                </div>
            )}
            <div className={`${styles.grid} ${columnsClass}`}>
                {items.map(item => (
                    <LinkTile
                        key={item.id}
                        title={item.title}
                        url={item.url}
                        icon={item.icon}
                        openInNewTab={props.openInNewTab}
                    />
                ))}
            </div>
        </div>
    );
};
