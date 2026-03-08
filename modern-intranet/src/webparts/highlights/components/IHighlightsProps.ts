/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IHighlightsProps {
    siteUrl: string;
    listId: string;
    titleColumn: string;
    descriptionColumn: string;
    bannerImageColumn: string;
    linkColumn: string;
    pinnedColumn?: string;
    maxItems: number;
    columns: number;
    showTitle: boolean;
    title: string;
    showBackgroundBar: boolean;
    titleBarStyle: 'solid' | 'underline';
    siteId: string;
    webId: string;
    context: WebPartContext;
}
