import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * Interface for SharePoint list information
 */
export interface IListInfo {
  id: string;
  title: string;
}

/**
 * Interface for SharePoint column information
 */
export interface IColumnInfo {
  id: string;
  internalName: string;
  title: string;
  type: string;
}

/**
 * Interface for attachment information
 */
export interface IAttachment {
  fileName: string;
  url: string;
}

/**
 * Interface for FAQ item
 */
export interface IFaqItem {
  id: number;
  title: string;
  description: string;
  category: string;
  subCategory?: string;
  isTopFaq: boolean;
  attachments: IAttachment[];
}

/**
 * SharePoint Service - Helper methods for retrieving lists, columns, and items
 */
export class SharePointService {
  private _context: WebPartContext;
  private _siteUrl: string;

  constructor(context: WebPartContext) {
    this._context = context;
    this._siteUrl = context.pageContext.web.absoluteUrl;
  }

  /**
   * Get all lists from the current site
   * Filters out hidden lists and system lists
   */
  public getLists(): Promise<IListInfo[]> {
    const endpoint = this._siteUrl + '/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 100&$select=Id,Title&$orderby=Title';

    return this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error('Error fetching lists: ' + response.statusText);
        }
        return response.json();
      })
      .then((data: { value: Array<{ Id: string; Title: string }> }) => {
        const lists: IListInfo[] = [];
        let i = 0;
        while (i < data.value.length) {
          lists.push({
            id: data.value[i].Id,
            title: data.value[i].Title
          });
          i++;
        }
        return lists;
      });
  }

  /**
   * Get columns for a specific list
   * Returns Text, Note, Choice, and Boolean columns
   */
  public getListColumns(listId: string): Promise<IColumnInfo[]> {
    // Filter for Text, Note, Choice, and Boolean field types
    const endpoint = this._siteUrl + '/_api/web/lists(guid\'' + listId + '\')/fields?$filter=(TypeAsString eq \'Text\' or TypeAsString eq \'Note\' or TypeAsString eq \'Choice\' or TypeAsString eq \'Boolean\') and Hidden eq false and ReadOnlyField eq false&$select=Id,InternalName,Title,TypeAsString&$orderby=Title';

    return this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error('Error fetching columns: ' + response.statusText);
        }
        return response.json();
      })
      .then((data: { value: Array<{ Id: string; InternalName: string; Title: string; TypeAsString: string }> }) => {
        const columns: IColumnInfo[] = [];
        let i = 0;
        while (i < data.value.length) {
          columns.push({
            id: data.value[i].Id,
            internalName: data.value[i].InternalName,
            title: data.value[i].Title,
            type: data.value[i].TypeAsString
          });
          i++;
        }
        return columns;
      });
  }

  /**
   * Get all items from a list with specified columns
   * Also fetches attachments for each item
   */
  public getListItems(
    listId: string,
    titleColumn: string,
    descriptionColumn: string,
    categoryColumn?: string,
    subCategoryColumn?: string,
    topFaqColumn?: string
  ): Promise<IFaqItem[]> {
    // Build select query - always include Id and Attachments
    const selectFields: string[] = ['Id', titleColumn];
    if (descriptionColumn !== titleColumn) {
      selectFields.push(descriptionColumn);
    }
    if (categoryColumn && categoryColumn !== titleColumn && categoryColumn !== descriptionColumn) {
      selectFields.push(categoryColumn);
    }
    if (subCategoryColumn && subCategoryColumn !== titleColumn && subCategoryColumn !== descriptionColumn && subCategoryColumn !== categoryColumn) {
      selectFields.push(subCategoryColumn);
    }
    if (topFaqColumn && selectFields.indexOf(topFaqColumn) === -1) {
      selectFields.push(topFaqColumn);
    }
    selectFields.push('Attachments');

    const endpoint = this._siteUrl + '/_api/web/lists(guid\'' + listId + '\')/items?$select=' + selectFields.join(',') + '&$top=500';

    const self = this;

    return this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error('Error fetching items: ' + response.statusText);
        }
        return response.json();
      })
      .then((data: { value: Array<{ [key: string]: unknown }> }) => {
        // Process items and fetch attachments
        const itemPromises: Array<Promise<IFaqItem>> = [];

        let i = 0;
        while (i < data.value.length) {
          const item = data.value[i];
          itemPromises.push(self._processItem(item, listId, titleColumn, descriptionColumn, categoryColumn, topFaqColumn, subCategoryColumn));
          i++;
        }

        return Promise.all(itemPromises);
      })
      .then((items: IFaqItem[]) => {
        // Filter out items without titles
        const filteredItems: IFaqItem[] = [];
        let i = 0;
        while (i < items.length) {
          if (items[i].title && items[i].title.trim() !== '') {
            filteredItems.push(items[i]);
          }
          i++;
        }
        return filteredItems;
      });
  }

  /**
   * Process a single item and fetch its attachments
   */
  private _processItem(
    item: { [key: string]: unknown },
    listId: string,
    titleColumn: string,
    descriptionColumn: string,
    categoryColumn?: string,
    topFaqColumn?: string,
    subCategoryColumn?: string
  ): Promise<IFaqItem> {
    const self = this;

    // Determine if item is marked as Top FAQ
    let isTopFaq = false;
    if (topFaqColumn && item[topFaqColumn] !== undefined) {
      const topFaqValue = item[topFaqColumn];
      // Handle Boolean (Yes/No) columns
      if (typeof topFaqValue === 'boolean') {
        isTopFaq = topFaqValue === true;
      }
      // Handle Choice columns with "Yes"/"No" values
      else if (typeof topFaqValue === 'string') {
        const lowerValue = topFaqValue.trim().toLowerCase();
        isTopFaq = lowerValue === 'yes' || lowerValue === 'true' || lowerValue === '1' || lowerValue === 'y';
      }
      // Handle numeric values (1 = yes)
      else if (typeof topFaqValue === 'number') {
        isTopFaq = topFaqValue === 1;
      }
    }

    const faqItem: IFaqItem = {
      id: item.Id as number,
      title: (item[titleColumn] as string) || '',
      description: (item[descriptionColumn] as string) || '',
      category: categoryColumn ? ((item[categoryColumn] as string) || '') : '',
      subCategory: subCategoryColumn ? ((item[subCategoryColumn] as string) || '') : undefined,
      isTopFaq: isTopFaq,
      attachments: []
    };

    // If item has attachments, fetch them
    if (item.Attachments) {
      return self._getItemAttachments(listId, item.Id as number)
        .then((attachments: IAttachment[]) => {
          faqItem.attachments = attachments;
          return faqItem;
        });
    }

    return Promise.resolve(faqItem);
  }

  /**
   * Get attachments for a specific list item
   */
  private _getItemAttachments(listId: string, itemId: number): Promise<IAttachment[]> {
    const endpoint = this._siteUrl + '/_api/web/lists(guid\'' + listId + '\')/items(' + itemId + ')/AttachmentFiles';
    const self = this;

    return this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          console.warn('Could not fetch attachments for item ' + itemId);
          return { value: [] };
        }
        return response.json();
      })
      .then((data: { value: Array<{ FileName: string; ServerRelativeUrl: string }> }) => {
        const attachments: IAttachment[] = [];
        const baseUrl = self._context.pageContext.web.absoluteUrl.split('/').slice(0, 3).join('/');

        let i = 0;
        while (i < data.value.length) {
          attachments.push({
            fileName: data.value[i].FileName,
            url: baseUrl + data.value[i].ServerRelativeUrl
          });
          i++;
        }

        return attachments;
      })
      .catch((error: Error) => {
        console.warn('Error fetching attachments:', error);
        return [];
      });
  }

  /**
   * Get a specific list by ID
   */
  public getListById(listId: string): Promise<IListInfo | null> {
    const endpoint = this._siteUrl + '/_api/web/lists(guid\'' + listId + '\')?$select=Id,Title';

    return this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return null;
        }
        return response.json();
      })
      .then((data: { Id: string; Title: string } | null) => {
        if (!data) {
          return null;
        }
        return {
          id: data.Id,
          title: data.Title
        };
      })
      .catch(() => {
        return null;
      });
  }

  /**
   * Check if a column exists in a list
   */
  public columnExists(listId: string, columnInternalName: string): Promise<boolean> {
    const endpoint = this._siteUrl + '/_api/web/lists(guid\'' + listId + '\')/fields?$filter=InternalName eq \'' + columnInternalName + '\'&$select=Id';

    return this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return false;
        }
        return response.json();
      })
      .then((data: { value: Array<{ Id: string }> }) => {
        return data.value && data.value.length > 0;
      })
      .catch(() => {
        return false;
      });
  }
}
