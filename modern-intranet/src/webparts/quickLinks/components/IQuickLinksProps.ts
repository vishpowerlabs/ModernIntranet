import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IQuickLinksProps {
    siteUrl: string;
    listId: string;
    titleColumn: string;
    linkColumn: string;
    iconColumn: string;
    pinnedColumn?: string;
    columnsPerRow: number;
    openInNewTab: boolean;
    showTitle?: boolean;
    title?: string;
    showBackgroundBar?: boolean;
    context: WebPartContext;
}
