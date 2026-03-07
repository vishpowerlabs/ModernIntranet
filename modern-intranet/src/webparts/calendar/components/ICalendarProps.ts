/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ICalendarEvent {
    id: string;
    title: string;
    start: Date;
    end: Date;
    location?: string;
    isAllDay?: boolean;
}

export interface ICalendarProps {
    context: WebPartContext;
    siteUrl: string;
    listId: string;
    titleColumn: string;
    dateColumn: string;
    endDateColumn: string;
    locationColumn: string;
    defaultView: 'day' | 'week' | 'month' | 'year';
    yearViewType: 'grid' | 'timeline';
    showTitle: boolean;
    title: string;
    showBackgroundBar: boolean;
}
