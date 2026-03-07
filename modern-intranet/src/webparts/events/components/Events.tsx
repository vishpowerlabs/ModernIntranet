import * as React from 'react';
import { useState, useEffect } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './Events.module.scss';
import { IEventsProps } from './IEventsProps';
import { EventCard } from './EventCard';
import { EmptyState } from '../../../common/components/EmptyState/EmptyState';

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

interface IEventItem {
  id: string;
  title: string;
  date: Date;
  imageUrl: string;
  location: string;
  linkUrl: string;
  pinned: boolean;
}

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

  return '';
};

export const Events: React.FC<IEventsProps> = (props) => {
  const [items, setItems] = useState<IEventItem[]>([]);
  const [loading, setLoading] = useState<boolean>(true);

  useEffect(() => {
    const fetchItems = async (): Promise<void> => {
      if (!props.listId || !props.titleColumn || !props.dateColumn) {
        setLoading(false);
        return;
      }

      try {
        setLoading(true);
        const today = new Date().toISOString();
        
        // Select logic
        const selectCols = [
          props.titleColumn,
          props.dateColumn,
          props.activeColumn,
          props.pinnedColumn,
          props.imageColumn,
          props.linkColumn,
          props.locationColumn,
          'Id',
          'FileDirRef'
        ].filter(v => !!v).join(',');

        let filter = `${props.dateColumn} ge datetime'${today}'`;
        if (props.activeColumn) {
          filter += ` and ${props.activeColumn} eq 1`;
        }

        const listUrl = `${props.siteUrl}/_api/web/lists(guid'${props.listId}')/items?$select=${selectCols}&$filter=${filter}&$orderby=${props.dateColumn} asc&$top=${props.maxItems}`;

        const response: SPHttpClientResponse = await props.context.spHttpClient.get(
          listUrl,
          SPHttpClient.configurations.v1
        );

        if (response.ok) {
          const data = await response.json();
          const formattedItems: IEventItem[] = data.value.map((row: ISharePointRow) => {
            const imageUrl = props.imageColumn ? getImageUrl(row, props.imageColumn, props.siteUrl, props.siteId, props.webId) : '';

            const linkData = props.linkColumn ? row[props.linkColumn] : '';
            let linkUrl = '';
            if (linkData) {
              linkUrl = (linkData as { Url?: string }).Url || String(linkData);
            }

            return {
              id: String(row.Id),
              title: String(row[props.titleColumn] || ''),
              date: new Date(String(row[props.dateColumn])),
              imageUrl,
              location: props.locationColumn ? String(row[props.locationColumn] || '') : '',
              linkUrl,
              pinned: props.pinnedColumn ? !!row[props.pinnedColumn] : false
            };
          }).sort((a: IEventItem, b: IEventItem) => {
            if (a.pinned !== b.pinned) {
              return a.pinned ? -1 : 1;
            }
            return a.date.getTime() - b.date.getTime();
          });
          setItems(formattedItems);
        }
      } catch (error) {
        console.error("Error fetching events:", error);
      } finally {
        setLoading(false);
      }
    };

    fetchItems().catch(err => {
      console.error("Error in fetchItems:", err);
    });
  }, [props.siteUrl, props.listId, props.titleColumn, props.dateColumn, props.activeColumn, props.pinnedColumn, props.imageColumn, props.linkColumn, props.locationColumn, props.maxItems, props.siteId, props.webId]);

  if (loading) {
    return null;
  }

  return (
    <section className={styles.eventsContainer}>
      {(props.showTitle && props.title) || (props.showViewAll && props.viewAllUrl) ? (
        <div className={styles.webpartHeader}>
          {props.showTitle && props.title && (
            <div className={styles.titleContainer}>
              <h2>{props.title}</h2>
              {props.showBackgroundBar && <div className={styles.backgroundBar} />}
            </div>
          )}
          {props.showViewAll && props.viewAllUrl && (
            <a href={props.viewAllUrl} className={styles.viewAll}>View All</a>
          )}
        </div>
      ) : null}

      {items.length === 0 ? (
        <EmptyState
          icon="Calendar"
          message="No upcoming events found."
        />
      ) : (
        <div className={`${styles.eventsGrid} ${(styles as any)[`cols${props.itemsPerRow}`]}`}>
          {items.map(item => (
            <EventCard
              key={item.id}
              title={item.title}
              date={item.date}
              imageUrl={item.imageUrl}
              location={item.location}
              linkUrl={item.linkUrl}
            />
          ))}
        </div>
      )}
    </section>
  );
};
