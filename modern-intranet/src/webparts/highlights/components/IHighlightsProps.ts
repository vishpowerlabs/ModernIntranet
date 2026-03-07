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
    siteId: string;
    webId: string;
    context: WebPartContext;
}
