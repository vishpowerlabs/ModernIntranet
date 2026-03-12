/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { useState, useEffect } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import styles from './Highlights.module.scss';
import { IHighlightsProps } from './IHighlightsProps';
import { HighlightCard } from './HighlightCard';
import { EmptyState } from '../../../common/components/EmptyState/EmptyState';
import { WebPartHeader } from '../../../common/components/WebPartHeader/WebPartHeader';

interface ISharePointImageMetadata {
    fileName?: string;
    serverRelativeUrl?: string;
    serverUrl?: string;
    Url?: string;
}

interface ISharePointRow {
    Id?: string;
    FileDirRef?: string;
    [key: string]: string | number | boolean | undefined | null;
}

interface IHighlightItem {
    id: string;
    title: string;
    description: string;
    imageUrl: string;
    linkUrl: string;
    pinned: boolean;
    created: Date;
}

const getPreviewUrl = (fileName: string, siteUrl: string, siteId: string, webId: string): string => {
    const match = /([a-f\d-]{32,36})/i.exec(fileName);
    if (!match) return '';

    let guid = match[1];
    if (guid.length === 32) {
        guid = `${guid.slice(0, 8)}-${guid.slice(8, 12)}-${guid.slice(12, 16)}-${guid.slice(16, 20)}-${guid.slice(20)}`;
    }
    return `${siteUrl}/_layouts/15/getpreview.ashx?guidSite=${siteId}&guidWeb=${webId}&guidFile=${guid}&clientType=image`;
};

const resolveImageObject = (rawValue: string | object): ISharePointImageMetadata | null => {
    if (typeof rawValue === 'string' && rawValue.startsWith('{')) {
        try { return JSON.parse(rawValue); } catch { return null; }
    }
    return typeof rawValue === 'object' ? rawValue : null;
};

const getImageUrl = (rowItem: ISharePointRow, imageColumn: string, siteUrl: string, siteId: string, webId: string): string => {
    const rawValue = rowItem[imageColumn];
    if (!rawValue) return '';

    const imageObj = resolveImageObject(rawValue as string | object);
    if (!imageObj) return typeof rawValue === 'string' ? rawValue : '';

    const url = imageObj.serverRelativeUrl || imageObj.serverUrl || imageObj.Url;
    if (url) {
        return (!url.startsWith('http') && url.startsWith('/'))
            ? `${new URL(siteUrl).origin}${url}`
            : url;
    }

    if (imageObj.fileName && rowItem.FileDirRef && rowItem.Id) {
        const origin = new URL(siteUrl).origin;
        return `${origin}${rowItem.FileDirRef}/Attachments/${rowItem.Id}/${imageObj.fileName}`;
    }

    return imageObj.fileName ? getPreviewUrl(imageObj.fileName, siteUrl, siteId, webId) : '';
};

export const Highlights: React.FC<IHighlightsProps> = (props) => {
    const [items, setItems] = useState<IHighlightItem[]>([]);
    const [loading, setLoading] = useState<boolean>(true);

    useEffect(() => {
        const fetchItems = async (): Promise<void> => {
            if (!props.siteUrl || !props.listId || !props.titleColumn || !props.bannerImageColumn || !props.linkColumn) {
                setLoading(false);
                return;
            }

            try {
                setLoading(true);
                const selectCols = [
                    props.titleColumn,
                    props.descriptionColumn,
                    props.bannerImageColumn,
                    props.linkColumn,
                    props.pinnedColumn,
                    'Id',
                    'FileDirRef',
                    'Created'
                ].filter(v => !!v).join(',');

                const listUrl = `${props.siteUrl}/_api/web/lists(guid'${props.listId}')/items?$select=${selectCols}&$orderby=Created desc&$top=${props.maxItems}`;

                const response: SPHttpClientResponse = await props.context.spHttpClient.get(
                    listUrl,
                    SPHttpClient.configurations.v1
                );

                if (response.ok) {
                    const data = await response.json();
                    const formattedItems: IHighlightItem[] = data.value.map((row: ISharePointRow) => {
                        const imageUrl = getImageUrl(row, props.bannerImageColumn, props.siteUrl, props.siteId, props.webId);

                        const linkData = row[props.linkColumn];
                        let linkUrl = '';
                        if (linkData) {
                            linkUrl = (linkData as { Url?: string }).Url || String(linkData);
                        }

                        return {
                            id: String(row.Id),
                            title: String(row[props.titleColumn] || ''),
                            description: String(row[props.descriptionColumn] || ''),
                            imageUrl,
                            linkUrl,
                            pinned: props.pinnedColumn ? !!row[props.pinnedColumn] : false,
                            created: new Date(String(row.Created))
                        };
                    }).sort((a: IHighlightItem, b: IHighlightItem) => {
                        if (a.pinned !== b.pinned) {
                            return a.pinned ? -1 : 1;
                        }
                        return b.created.getTime() - a.created.getTime();
                    });
                    setItems(formattedItems);
                }
            } catch (error) {
                console.error("Error fetching highlights:", error);
            } finally {
                setLoading(false);
            }
        };

        fetchItems().catch(err => {
            console.error("Error in fetchItems:", err);
        });
    }, [props.siteUrl, props.listId, props.titleColumn, props.descriptionColumn, props.bannerImageColumn, props.linkColumn, props.pinnedColumn, props.maxItems, props.siteId, props.webId]);

    const renderHeader = (): JSX.Element => (
        <WebPartHeader
            title={props.title || ''}
            showTitle={!!props.showTitle}
            showBackgroundBar={!!props.showBackgroundBar}
            titleBarStyle={props.titleBarStyle || 'underline'}
        />
    );

    if (loading) {
        return (
            <section className={styles.highlightsContainer}>
                {renderHeader()}
                <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '200px' }}>
                    <Spinner size={SpinnerSize.large} label="Loading highlights..." />
                </div>
            </section>
        );
    }

    const isConfigured = props.siteUrl && props.listId && props.titleColumn && props.bannerImageColumn && props.linkColumn;

    if (!isConfigured) {
        return (
            <section className={styles.highlightsContainer}>
                <EmptyState
                    icon="Highlight"
                    title="Highlights - Configuration Required"
                    message="Please complete the web part configuration to display content."
                    description="You need to specify the Site URL, List ID, and map the required columns (Title, Banner Image, and Link) in the property pane."
                />
            </section>
        );
    }

    if (items.length === 0) {
        return (
            <section className={styles.highlightsContainer}>
                {renderHeader()}
                <EmptyState
                    icon="Highlight"
                    title="No Highlights Found"
                    message="There are no items to display from the selected list."
                    description="Add items to your SharePoint list or check your filter settings if applicable."
                />
            </section>
        );
    }

    return (
        <section className={styles.highlightsContainer}>
            {renderHeader()}
            <div className={`${styles.highlightsGrid} ${props.columns === 2 ? styles.cols2 : styles.cols3}`}>
                {items.map(item => {
                    const colSize = props.columns === 2 ? 6 : 4;
                    const colClass = `ms-sm12 ms-md${colSize} ms-lg${colSize}`;
                    return (
                        <div key={item.id} className={`${styles.highlightsCol} ${colClass}`}>
                            <HighlightCard
                                title={item.title}
                                description={item.description}
                                imageUrl={item.imageUrl}
                                linkUrl={item.linkUrl}
                            />
                        </div>
                    );
                })}
            </div>
        </section>
    );
};
