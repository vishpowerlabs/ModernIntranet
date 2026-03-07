/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import { ISite, IListInfo, IColumnInfo, IDocument } from '../models';

export interface ISiteListService {
    getSites(): Promise<ISite[]>;
    getLists(siteUrl: string): Promise<IListInfo[]>;
    getColumns(siteUrl: string, listId: string, typeFilter?: string): Promise<IColumnInfo[]>;
    getDocuments(
        siteUrl: string,
        listId: string,
        categoryField?: string,
        subCategoryField?: string,
        descriptionField?: string,
        pinnedField?: string
    ): Promise<IDocument[]>;
}
