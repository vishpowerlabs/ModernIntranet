import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Guid } from "@microsoft/sp-core-library";
import { IDocument } from "../models/interfaces";
import { ISPItemsResponse } from "../models/ISPData";

export class SPService {
    private static _instance: SPService;
    private context!: WebPartContext;

    private constructor() {
        // Private constructor
    }

    public static get Instance(): SPService {
        if (!this._instance) {
            this._instance = new SPService();
        }
        return this._instance;
    }

    public setup(context: WebPartContext): void {
        this.context = context;
    }

    public async getDocuments(
        listId: string,
        categoryField: string = 'Category',
        subCategoryField: string = 'SubCategory',
        descriptionField: string = 'Description',
        pinnedField?: string,
        siteUrl?: string
    ): Promise<IDocument[]> {
        if (!this.context || !listId) return [];

        const webUrl = siteUrl || this.context.pageContext.web.absoluteUrl;

        try {
            // Build Select fields
            const selectFields = [
                'Id', 'Title', 'Modified', 'Created',
                'File/Name', 'File/ServerRelativeUrl', 'File/UniqueId',
                'Editor/Title', 'Author/Title',
                categoryField
            ];

            if (subCategoryField) selectFields.push(subCategoryField);
            if (descriptionField) selectFields.push(descriptionField);
            if (pinnedField) selectFields.push(pinnedField);

            // Remove duplicates just in case
            const uniqueSelects = Array.from(new Set(selectFields)).join(',');

            const endpoint = `${webUrl}/_api/web/lists(guid'${listId}')/items?$select=${uniqueSelects}&$expand=File,Editor,Author`;

            const response: SPHttpClientResponse = await this.context.spHttpClient.get(
                endpoint,
                SPHttpClient.configurations.v1
            );

            if (!response.ok) {
                throw new Error(`Error fetching documents: ${response.statusText}`);
            }

            const data = await response.json();

            return data.value.map((item: any) => ({
                Id: item.Id,
                Title: item.Title || item.File?.Name || 'Untitled',
                Name: item.File?.Name || 'Untitled',
                Description: item[descriptionField] || '',
                FileRef: item.File?.ServerRelativeUrl,
                Pinned: pinnedField ? !!item[pinnedField] : false, // Map pinned field
                Modified: item.Modified,
                Created: item.Created,
                Category: item[categoryField] || 'Uncategorized',
                SubCategory: subCategoryField ? (item[subCategoryField] || '') : '',
                Author: item.Author?.Title,
                Editor: item.Editor?.Title,
                UniqueId: item.File?.UniqueId,
                ServerRelativeUrl: item.File?.ServerRelativeUrl
            }));
        } catch (error) {
            console.error('Error fetching documents', error);
            return [];
        }
    }

    public async logAccessRequest(
        listId: string,
        email: string,
        fileId: string,
        emailField: string,
        fileIdField: string,
        requestIdField: string,
        dateField: string,
        siteUrl?: string
    ): Promise<{ status: string, itemId?: number }> {
        if (!this.context || !listId) {
            throw new Error("Service not initialized or List ID missing");
        }

        const webUrl = siteUrl || this.context.pageContext.web.absoluteUrl;

        const checkEndpoint = `${webUrl}/_api/web/lists(guid'${listId}')/items?$select=Id,${emailField},${fileIdField}&$filter=${emailField} eq '${email}' and ${fileIdField} eq '${fileId}'`;

        try {
            const response = await this.context.spHttpClient.get(checkEndpoint, SPHttpClient.configurations.v1);
            const data: ISPItemsResponse = await response.json();

            if (data.value && data.value.length > 0) {
                return { status: 'Exists', itemId: data.value[0].Id };
            } else {
                // Generate Request ID (simple random GUID-like string or use sp-core-library Guid)
                // Generate Request ID (simple random GUID-like string or use sp-core-library Guid)
                const requestId = Guid.newGuid().toString();
                const requestDate = new Date().toISOString();

                const body: any = {
                    Title: `Request: ${email} - ${fileId}`
                };
                body[emailField] = email;
                body[fileIdField] = fileId;
                body[requestIdField] = requestId;
                body[dateField] = requestDate;

                await this.context.spHttpClient.post(
                    `${webUrl}/_api/web/lists(guid'${listId}')/items`,
                    SPHttpClient.configurations.v1,
                    {
                        headers: {
                            'Accept': 'application/json;odata=nometadata',
                            'Content-type': 'application/json;odata=nometadata',
                            'odata-version': ''
                        },
                        body: JSON.stringify(body)
                    }
                );
                return { status: 'Created' };
            }
        } catch (error) {
            console.error("Error logging access request", error);
            throw error;
        }
    }

    public async setRequestReminder(listId: string, itemId: number, reminderField: string, siteUrl?: string): Promise<void> {
        if (!this.context || !listId || !itemId) {
            throw new Error("Service not initialized or arguments missing");
        }

        const body: any = {};
        body[reminderField] = 'Yes'; // Assuming Text field or compatible string for now

        const webUrl = siteUrl || this.context.pageContext.web.absoluteUrl;

        try {
            await this.context.spHttpClient.post(
                `${webUrl}/_api/web/lists(guid'${listId}')/items(${itemId})`,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'Content-type': 'application/json;odata=nometadata',
                        'X-HTTP-Method': 'MERGE',
                        'IF-MATCH': '*'
                    },
                    body: JSON.stringify(body)
                }
            );
        } catch (error) {
            console.error("Error setting reminder", error);
            throw error;
        }
    }


}
