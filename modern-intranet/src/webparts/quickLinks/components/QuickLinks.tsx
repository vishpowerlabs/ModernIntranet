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

                if (rawLink && typeof rawLink === 'object' && (rawLink as ISharePointLinkValue).Url) {
                    urlValue = (rawLink as ISharePointLinkValue).Url || '';
                } else {
                    urlValue = (rawLink as string) || '';
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

    if (!props.siteUrl || !props.listId) {
        return <EmptyState icon="Link" message="Please configure the web part to select a data source." />;
    }

    if (error) {
        return <EmptyState icon="Error" message={error} />;
    }

    if (loading) {
        return <Spinner size={SpinnerSize.large} label="Loading links..." />;
    }

    if (items.length === 0) {
        return <EmptyState icon="SearchIssue" message="No visible links found in the selected list." />;
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
                <div className={styles.webpartHeader}>
                    <div className={styles.titleContainer}>
                        <h2>{props.title}</h2>
                        {props.showBackgroundBar && <div className={styles.backgroundBar} />}
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
