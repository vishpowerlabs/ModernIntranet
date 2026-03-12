/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

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
        if (!siteUrl || !listId) {
            console.warn('SiteListService.getDocuments: missing siteUrl or listId', { siteUrl, listId });
            return [];
        }

        const normalizedSiteUrl = siteUrl.endsWith('/') ? siteUrl.substring(0, siteUrl.length - 1) : siteUrl;

        try {
            const availableFields = await this._getAvailableFields(normalizedSiteUrl, listId);
            
            const titleExact = 'Title';
            const authorExact = 'Author';
            const editorExact = 'Editor';

            const getFieldName = (name: string): string => {
                if (!name) return '';
                return availableFields.get(name.toLowerCase()) || name;
            };

            const selectFields = ['Id', titleExact, 'Modified', 'Created', 'EncodedAbsUrl', 'FileLeafRef', 'FileRef'];
            const expands = [];

            if (availableFields.has('author')) {
                expands.push(authorExact);
                selectFields.push(`${authorExact}/Title`);
            }
            if (availableFields.has('editor')) {
                expands.push(editorExact);
                selectFields.push(`${editorExact}/Title`);
            }

            const catExact = categoryField ? getFieldName(categoryField) : '';
            if (catExact && !selectFields.includes(catExact)) selectFields.push(catExact);
            
            const subCatExact = subCategoryField ? getFieldName(subCategoryField) : '';
            if (subCatExact && !selectFields.includes(subCatExact)) selectFields.push(subCatExact);
            
            const descExact = descriptionField ? getFieldName(descriptionField) : '';
            if (descExact && !selectFields.includes(descExact)) selectFields.push(descExact);
            
            const pinExact = pinnedField ? getFieldName(pinnedField) : '';
            if (pinExact && !selectFields.includes(pinExact)) selectFields.push(pinExact);

            return await this._fetchTieredDocuments(normalizedSiteUrl, listId, selectFields, expands, {
                descExact,
                catExact,
                subCatExact,
                pinExact: pinExact || '',
                titleExact,
                authorExact,
                editorExact
            });

        } catch (error) {
            console.error('SiteListService.getDocuments caught error:', error);
            throw error;
        }
    }

    private async _getAvailableFields(siteUrl: string, listId: string): Promise<Map<string, string>> {
        const fieldsEndpoint = `${siteUrl}/_api/web/lists(guid'${listId}')/fields?$select=InternalName,StaticName,Title,TypeAsString`;
        const fieldsResponse = await this._context.spHttpClient.get(fieldsEndpoint, SPHttpClient.configurations.v1);
        
        const availableFields: Map<string, string> = new Map();
        if (fieldsResponse.ok) {
            const fieldsData = await fieldsResponse.json();
            (fieldsData.value || []).forEach((f: any) => {
                if (f.InternalName) availableFields.set(f.InternalName.toLowerCase(), f.InternalName);
                if (f.StaticName) availableFields.set(f.StaticName.toLowerCase(), f.InternalName);
            });
            console.log(`SiteListService: Identified ${availableFields.size} internal mappings.`);
        } else {
            console.warn(`SiteListService: Field discovery failed (${fieldsResponse.status}). Using system defaults.`);
        }
        return availableFields;
    }

    private async _fetchTieredDocuments(
        siteUrl: string, 
        listId: string, 
        selectFields: string[], 
        expands: string[],
        fieldNames: any
    ): Promise<IDocument[]> {
        const { descExact, catExact, subCatExact, pinExact, titleExact, authorExact, editorExact } = fieldNames;

        // TIER 1: Full Fetch
        const primarySelects = [...selectFields, 'File/Name', 'File/ServerRelativeUrl', 'File/UniqueId'];
        const primaryExpands = [...expands, 'File'];
        const primaryEndpoint = `${siteUrl}/_api/web/lists(guid'${listId}')/items?$select=${primarySelects.join(',')}&$expand=${primaryExpands.join(',')}`;

        const response = await this._context.spHttpClient.get(primaryEndpoint, SPHttpClient.configurations.v1);
        if (response.ok) {
            const data = await response.json();
            return this._mapDocuments(data.value || [], descExact, catExact, subCatExact, pinExact, titleExact, authorExact, editorExact);
        }

        // TIER 2: No File
        const tier2Endpoint = `${siteUrl}/_api/web/lists(guid'${listId}')/items?$select=${selectFields.join(',')}&$expand=${expands.join(',')}`;
        const t2Response = await this._context.spHttpClient.get(tier2Endpoint, SPHttpClient.configurations.v1);
        if (t2Response.ok) {
            const data = await t2Response.json();
            return this._mapDocuments(data.value || [], descExact, catExact, subCatExact, pinExact, titleExact, authorExact, editorExact);
        }

        // TIER 3: Minimal
        const minimalSelects = ['Id', titleExact, 'Modified', 'EncodedAbsUrl'];
        [catExact, subCatExact, descExact].forEach(f => {
            if (f && !minimalSelects.includes(f)) minimalSelects.push(f);
        });

        const tier3Endpoint = `${siteUrl}/_api/web/lists(guid'${listId}')/items?$select=${minimalSelects.join(',')}`;
        const t3Response = await this._context.spHttpClient.get(tier3Endpoint, SPHttpClient.configurations.v1);
        if (t3Response.ok) {
            const data = await t3Response.json();
            return this._mapDocuments(data.value || [], descExact, catExact, subCatExact, pinExact, titleExact, authorExact, editorExact);
        }

        throw new Error(`Critical OData Error. Status ${t3Response.status}.`);
    }

    private _mapDocuments(
        items: any[], 
        descriptionField: string, 
        categoryField: string, 
        subCategoryField: string, 
        pinnedField: string,
        titleField: string = 'Title',
        authorField: string = 'Author',
        editorField: string = 'Editor'
    ): IDocument[] {
        return items.map((item: any) => {
            const getVal = (field: string): string => {
                if (!field || !item[field]) return '';
                const val = item[field];
                if (typeof val === 'string') return val;
                if (typeof val === 'object') {
                    if (val.results && Array.isArray(val.results) && val.results.length > 0) {
                        return val.results.map((r: any) => r.Label || r.Title || r.Value || String(r)).join(', ');
                    }
                    return val.Label || val.Title || val.Value || val.TermGuid || JSON.stringify(val);
                }
                return String(val);
            };

            const isPinned = (field: string): boolean => {
                if (!field || !item[field]) return false;
                const val = item[field];
                return val === true || val === 'Yes' || String(val).toLowerCase() === 'true' || (typeof val === 'object' && (val.Value === 'Yes' || val.Value === true));
            };

            return {
                Id: item.Id,
                Title: item[titleField] || item.FileLeafRef || item.File?.Name || 'Untitled',
                Name: item.FileLeafRef || item.File?.Name || 'Untitled',
                Description: getVal(descriptionField),
                FileRef: item.FileRef || item.File?.ServerRelativeUrl,
                Pinned: isPinned(pinnedField),
                Modified: item.Modified,
                Created: item.Created,
                Category: getVal(categoryField) || 'Uncategorized',
                SubCategory: getVal(subCategoryField),
                Author: item[authorField]?.Title || '',
                Editor: item[editorField]?.Title || '',
                UniqueId: item.File?.UniqueId || `item-${item.Id}`,
                ServerRelativeUrl: item.FileRef || item.File?.ServerRelativeUrl,
                EncodedAbsUrl: item.EncodedAbsUrl || item.FileRef
            };
        });
    }
}
