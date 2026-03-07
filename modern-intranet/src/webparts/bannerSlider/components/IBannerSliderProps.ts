/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IBannerSliderProps {
    siteUrl: string;
    siteId: string;
    webId: string;
    listId: string;
    titleColumn: string;
    descriptionColumn: string;
    imageColumn: string;
    activeColumn: string;
    buttonTextColumn: string;
    pageLinkColumn: string;
    autoRotateInterval: number;
    showCta: boolean;
    showTitle?: boolean;
    title?: string;
    showBackgroundBar?: boolean;
    context: WebPartContext;
}
