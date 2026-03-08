/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IModernDocumentViewerProps {
    siteUrl: string;
    listId: string;
    categoryField: string;
    subCategoryField: string;
    descriptionField: string;
    pinnedField: string;
    enableSubCategory: boolean;
    categoryDisplayType: 'side' | 'top';
    pageSize: number;
    webPartTitle: string;
    webPartTitleFontSize: string;
    webPartDescription: string;
    webPartDescriptionFontSize: string;
    headerOpacity: number;
    showBackgroundBar: boolean;
    titleBarStyle: 'solid' | 'underline';
    context: WebPartContext;
}
