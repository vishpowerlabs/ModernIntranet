/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { useState, useEffect } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Callout, DirectionalHint } from '@fluentui/react/lib/Callout';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './Events.module.scss';
import { IEventsProps } from './IEventsProps';
import { EventCard } from './EventCard';
import { EmptyState } from '../../../common/components/EmptyState/EmptyState';
import { WebPartHeader } from '../../../common/components/WebPartHeader/WebPartHeader';

interface ISharePointRow {
  Id?: string;
  [key: string]: string | number | boolean | undefined | null;
}

interface IEventItem {
  id: string;
  title: string;
  date: Date;
  location: string;
  description: string;
  linkUrl: string;
  pinned: boolean;
}

export const Events: React.FC<IEventsProps> = (props) => {
  const [items, setItems] = useState<IEventItem[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [activePopoverId, setActivePopoverId] = useState<string | null>(null);
  const [popoverTarget, setPopoverTarget] = useState<HTMLElement | null>(null);
  const [currentPage, setCurrentPage] = useState<number>(1);

  useEffect(() => {
    const fetchItems = async (): Promise<void> => {
      if (!props.listId || !props.titleColumn || !props.dateColumn) {
        setLoading(false);
        return;
      }

      try {
        setLoading(true);
        const today = new Date().toISOString();

        const selectCols = [
          props.titleColumn,
          props.dateColumn,
          props.activeColumn,
          props.pinnedColumn,
          props.linkColumn,
          props.locationColumn,
          props.descriptionColumn,
          'Id'
        ].filter(v => !!v).join(',');

        let filter = `${props.dateColumn} ge datetime'${today}'`;
        if (props.activeColumn) {
          filter += ` and ${props.activeColumn} eq 1`;
        }

        const topCount = props.showPagination ? 100 : props.maxItems;
        const listUrl = `${props.siteUrl}/_api/web/lists(guid'${props.listId}')/items?$select=${selectCols}&$filter=${filter}&$orderby=${props.dateColumn} asc&$top=${topCount}`;

        const response: SPHttpClientResponse = await props.context.spHttpClient.get(
          listUrl,
          SPHttpClient.configurations.v1
        );

        if (response.ok) {
          const data = await response.json();
          const formattedItems: IEventItem[] = (data.value || []).map((row: ISharePointRow) => {
            const linkData = props.linkColumn ? row[props.linkColumn] : '';
            let linkUrl = '';
            if (linkData) {
              linkUrl = (linkData as { Url?: string }).Url || String(linkData);
            }

            const rawPinned = props.pinnedColumn ? row[props.pinnedColumn] : null;
            const isPinned = rawPinned === true || rawPinned === 1 || String(rawPinned).toLowerCase() === 'true' || String(rawPinned) === '1';

            return {
              id: String(row.Id),
              title: String(row[props.titleColumn] || ''),
              date: new Date(String(row[props.dateColumn])),
              location: props.locationColumn ? String(row[props.locationColumn] || '') : '',
              description: props.descriptionColumn ? String(row[props.descriptionColumn] || '') : '',
              linkUrl,
              pinned: isPinned
            };
          }).sort((a: IEventItem, b: IEventItem) => {
            if (a.pinned !== b.pinned) return a.pinned ? -1 : 1;
            return a.date.getTime() - b.date.getTime();
          });
          setItems(formattedItems);
          setCurrentPage(1);
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
  }, [props.siteUrl, props.listId, props.titleColumn, props.dateColumn, props.activeColumn, props.pinnedColumn, props.linkColumn, props.locationColumn, props.descriptionColumn, props.maxItems, props.showPagination]);

  const handleInfoClick = (event: React.MouseEvent<HTMLElement>, id: string): void => {
    setPopoverTarget(event.currentTarget);
    setActivePopoverId(id);
  };

  const handlePopoverDismiss = (): void => {
    setActivePopoverId(null);
    setPopoverTarget(null);
  };

  if (loading) {
    return (
      <section className={styles.eventsContainer}>
        <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '150px' }}>
          <Spinner size={SpinnerSize.large} label="Loading events..." />
        </div>
      </section>
    );
  }

  const isConfigured = props.siteUrl && props.listId && props.titleColumn && props.dateColumn;

  if (!isConfigured) {
    return (
      <section className={styles.eventsContainer}>
        <EmptyState
          icon="Calendar"
          title="Events - Configuration Required"
          message="Please complete the web part configuration to display events."
          description="Map the required columns (Title and Date) in the property pane."
        />
      </section>
    );
  }

  const renderHeader = (): JSX.Element | null => {
    return (
      <WebPartHeader
        title={props.title}
        showTitle={props.showTitle}
        showBackgroundBar={props.showBackgroundBar}
        titleBarStyle={props.titleBarStyle}
        actionTitle={props.showViewAll ? "View All" : undefined}
        actionUrl={props.viewAllUrl}
      />
    );
  };

  const renderPopover = (): JSX.Element | null => {
    if (!activePopoverId || !popoverTarget) return null;

    const item = items.find(i => i.id === activePopoverId);
    if (!item) return null;

    const day = item.date.getDate();
    const month = item.date.toLocaleString('default', { month: 'short' });
    const year = item.date.getFullYear();
    const fullDate = item.date.toLocaleDateString('default', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
    const time = item.date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

    return (
      <Callout
        target={popoverTarget}
        onDismiss={handlePopoverDismiss}
        setInitialFocus
        directionalHint={DirectionalHint.rightTopEdge}
        isBeakVisible={true}
        beakWidth={10}
        gapSpace={10}
        layerProps={{ eventBubblingEnabled: true }}
      >
        <div className={styles.popover}>
          <div className={styles.popTop}>
            <div className={styles.popDateLbl}>{month} {day}, {year}</div>
            <div className={styles.popTitle}>{item.title}</div>
          </div>
          <div className={styles.popBody}>
            <div className={styles.popRow}>
              <Icon iconName="MapPin" className={styles.icon} />
              <div className={styles.rowContent}>
                <strong>Location</strong>
                <span>{item.location || 'No location provided'}</span>
              </div>
            </div>
            <div className={styles.popRow}>
              <Icon iconName="Calendar" className={styles.icon} />
              <div className={styles.rowContent}>
                <strong>Date & Time</strong>
                <span>{fullDate} at {time}</span>
              </div>
            </div>
            {item.description && (
              <div className={styles.popRow}>
                <Icon iconName="Info" className={styles.icon} />
                <div className={styles.rowContent}>
                  <strong>Description</strong>
                  <span>{item.description}</span>
                </div>
              </div>
            )}
            {item.linkUrl && (
              <a href={item.linkUrl} target="_blank" rel="noopener noreferrer" className={styles.popoverLink}>
                View Event Details
              </a>
            )}
          </div>
        </div>
      </Callout>
    );
  };

  const renderPagination = (totalPages: number): JSX.Element | null => {
    if (!props.showPagination || totalPages <= 1) return null;

    return (
      <div className={styles.pagination}>
        <button 
          className={styles.pageBtn} 
          disabled={currentPage === 1}
          onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
          type="button"
        >
          <Icon iconName="ChevronLeft" />
        </button>
        <span className={styles.pageInfo}>
          Page {currentPage} of {totalPages}
        </span>
        <button 
          className={styles.pageBtn} 
          disabled={currentPage === totalPages}
          onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
          type="button"
        >
          <Icon iconName="ChevronRight" />
        </button>
      </div>
    );
  };

  const renderEvents = (): JSX.Element => {
    if (items.length === 0) {
      return (
        <EmptyState
          icon="Calendar"
          title="No Upcoming Events"
          message="There are no upcoming events to display."
        />
      );
    }

    const maxItems = props.maxItems || 6;
    const totalPages = Math.ceil(items.length / maxItems);
    
    const currentItems = props.showPagination 
      ? items.slice((currentPage - 1) * maxItems, currentPage * maxItems)
      : items;

    const layout = props.layout || 'grid';
    
    let content: JSX.Element;

    if (layout === 'list') {
      content = (
        <div className={styles.eventsList}>
          {currentItems.map(item => (
            <EventCard
              key={item.id}
              id={item.id}
              title={item.title}
              date={item.date}
              location={item.location}
              description={item.description}
              linkUrl={item.linkUrl}
              layout="list"
              onInfoClick={handleInfoClick}
            />
          ))}
        </div>
      );
    } else if (layout === 'compact') {
      content = (
        <div className={styles.eventsCompact}>
          {currentItems.map(item => (
            <EventCard
              key={item.id}
              id={item.id}
              title={item.title}
              date={item.date}
              location={item.location}
              description={item.description}
              linkUrl={item.linkUrl}
              layout="compact"
              onInfoClick={handleInfoClick}
            />
          ))}
        </div>
      );
    } else {
      const itemsPerRow = props.itemsPerRow || 3;
      const gridClasses = `${styles.eventsGrid} ${styles[`cols${itemsPerRow}` as keyof typeof styles] || ''}`;

      content = (
        <div className={gridClasses}>
          {currentItems.map(item => (
            <EventCard
              key={item.id}
              id={item.id}
              title={item.title}
              date={item.date}
              location={item.location}
              description={item.description}
              linkUrl={item.linkUrl}
              layout="grid"
              onInfoClick={handleInfoClick}
            />
          ))}
        </div>
      );
    }

    return (
      <>
        {content}
        {renderPagination(totalPages)}
      </>
    );
  };

  return (
    <section className={styles.eventsContainer}>
      {renderHeader()}
      {renderEvents()}
      {renderPopover()}
    </section>
  );
};
