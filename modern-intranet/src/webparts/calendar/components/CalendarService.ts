/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ICalendarEvent } from './ICalendarProps';

export class CalendarService {
    private readonly _context: WebPartContext;

    constructor(context: WebPartContext) {
        this._context = context;
    }

    public async getSPListEvents(
        siteUrl: string,
        listId: string,
        viewStart: Date,
        viewEnd: Date,
        mappings: { title: string; date: string; endDate?: string; location?: string }
    ): Promise<ICalendarEvent[]> {
        if (!siteUrl || !listId || !mappings.title || !mappings.date) return [];

        const startISO = viewStart.toISOString();
        const endISO = viewEnd.toISOString();

        // Simple filter for SP REST API
        const filter = `${mappings.date} ge datetime'${startISO}' and ${mappings.date} le datetime'${endISO}'`;
        const selectCols = [`Id`, mappings.title, mappings.date];
        if (mappings.endDate) selectCols.push(mappings.endDate);
        if (mappings.location) selectCols.push(mappings.location);

        const select = selectCols.join(',');
        const endpoint = `${siteUrl}/_api/web/lists(guid'${listId}')/items?$filter=${filter}&$select=${select}&$orderby=${mappings.date} asc`;

        const response: SPHttpClientResponse = await this._context.spHttpClient.get(
            endpoint,
            SPHttpClient.configurations.v1
        );

        if (!response.ok) return [];

        const data = await response.json();
        return (data.value || []).map((item: any) => ({
            id: item.Id.toString(),
            title: item[mappings.title],
            start: new Date(item[mappings.date]),
            end: mappings.endDate && item[mappings.endDate] ? new Date(item[mappings.endDate]) : new Date(item[mappings.date]),
            location: mappings.location ? item[mappings.location] : undefined
        }));
    }
}
