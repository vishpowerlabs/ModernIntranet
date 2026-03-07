import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ISite, IListInfo, IColumnInfo, IDocument } from '../models';
import { ISiteListService } from './ISiteListService';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface ISearchResultRow {
    Cells: Array<{ Key: string; Value: string }>;
}

export class SiteListService implements ISiteListService {
    private readonly _context: WebPartContext;

    constructor(context: WebPartContext) {
        this._context = context;
    }

    public async getSites(): Promise<ISite[]> {
        try {
            const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/search/query?querytext='contentclass:STS_Site'&selectproperties='Title,Path'`;
            const response: SPHttpClientResponse = await this._context.spHttpClient.get(
                endpoint,
                SPHttpClient.configurations.v1
            );

            if (!response.ok) {
                throw new Error(`Error fetching sites: ${response.statusText}`);
            }

            const data = await response.json();
            const rows: ISearchResultRow[] = data?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];

            const sites: ISite[] = rows.map((row) => {
                const titleCell = row.Cells.find(c => c.Key === 'Title');
                const pathCell = row.Cells.find(c => c.Key === 'Path');
                return {
                    title: titleCell ? titleCell.Value : 'Unknown',
                    url: pathCell ? pathCell.Value : ''
                };
            });

            return sites;
        } catch (error) {
            console.error('SiteListService.getSites error:', error);
            return [];
        }
    }

    public async getLists(siteUrl: string): Promise<IListInfo[]> {
        if (!siteUrl) return [];

        try {
            const endpoint = `${siteUrl}/_api/web/lists?$filter=Hidden eq false&$select=Id,Title`;
            const response: SPHttpClientResponse = await this._context.spHttpClient.get(
                endpoint,
                SPHttpClient.configurations.v1
            );

            if (!response.ok) {
                throw new Error(`Error fetching lists: ${response.statusText}`);
            }

            const data = await response.json();
            return (data.value || []).map((list: { Id: string; Title: string }) => ({
                id: list.Id,
                title: list.Title
            }));
        } catch (err) {
            console.error('getLists error:', (err as Error).message);
            return [];
        }
    }

    public async getColumns(siteUrl: string, listId: string, typeFilter?: string): Promise<IColumnInfo[]> {
        if (!siteUrl || !listId) return [];

        try {
            const endpoint = `${siteUrl}/_api/web/lists(guid'${listId}')/fields?$filter=Hidden eq false&$select=InternalName,Title,TypeAsString`;
            const response: SPHttpClientResponse = await this._context.spHttpClient.get(
                endpoint,
                SPHttpClient.configurations.v1
            );

            if (!response.ok) {
                throw new Error(`Error fetching columns: ${response.statusText}`);
            }

            const data = await response.json();
            let columns = (data.value || []).map((col: { InternalName: string; Title: string; TypeAsString: string }) => ({
                internalName: col.InternalName,
                title: col.Title,
                typeAsString: col.TypeAsString
            }));

            if (typeFilter) {
                const filterSet = new Set(typeFilter.split(',').map(f => f.trim().toLowerCase()));
                columns = columns.filter((col: IColumnInfo) => filterSet.has(col.typeAsString.toLowerCase()));
            }

            return columns;
        } catch (err) {
            console.error('getColumns error:', (err as Error).message);
            return [];
        }
    }

    public async getDocuments(
        siteUrl: string,
        listId: string,
        categoryField: string = 'Category',
        subCategoryField: string = 'SubCategory',
        descriptionField: string = 'Description',
        pinnedField?: string
    ): Promise<IDocument[]> {
        if (!siteUrl || !listId) return [];

        try {
            // Build Select fields
            const selectFields = [
                'Id', 'Title', 'Modified', 'Created',
                'File/Name', 'File/ServerRelativeUrl', 'File/UniqueId',
                'Editor/Title', 'Author/Title', 'EncodedAbsUrl',
                categoryField
            ];

            if (subCategoryField) selectFields.push(subCategoryField);
            if (descriptionField) selectFields.push(descriptionField);
            if (pinnedField) selectFields.push(pinnedField);

            const uniqueSelects = Array.from(new Set(selectFields)).join(',');
            const endpoint = `${siteUrl}/_api/web/lists(guid'${listId}')/items?$select=${uniqueSelects}&$expand=File,Editor,Author`;

            const response: SPHttpClientResponse = await this._context.spHttpClient.get(
                endpoint,
                SPHttpClient.configurations.v1
            );

            if (!response.ok) {
                throw new Error(`Error fetching documents: ${response.statusText}`);
            }

            const data = await response.json();

            return (data.value || []).map((item: any) => ({
                Id: item.Id,
                Title: item.Title || item.File?.Name || 'Untitled',
                Name: item.File?.Name || 'Untitled',
                Description: item[descriptionField] || '',
                FileRef: item.File?.ServerRelativeUrl,
                Pinned: pinnedField ? (item[pinnedField] === true || item[pinnedField] === 'Yes' || String(item[pinnedField]).toLowerCase() === 'true') : false,
                Modified: item.Modified,
                Created: item.Created,
                Category: item[categoryField] || 'Uncategorized',
                SubCategory: subCategoryField ? (item[subCategoryField] || '') : '',
                Author: item.Author?.Title,
                Editor: item.Editor?.Title,
                UniqueId: item.File?.UniqueId,
                ServerRelativeUrl: item.File?.ServerRelativeUrl,
                EncodedAbsUrl: item.EncodedAbsUrl
            }));
        } catch (error) {
            console.error('SiteListService.getDocuments error:', error);
            return [];
        }
    }
}
