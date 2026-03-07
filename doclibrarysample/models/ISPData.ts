export interface ISPList {
    Id: string;
    Title: string;
}

export interface ISPField {
    InternalName: string;
    Title: string;
}

export interface ISPListItem {
    Id: number;
    [key: string]: any; // Allow dynamic fields
}

export interface ISPListsResponse {
    value: ISPList[];
}

export interface ISPFieldsResponse {
    value: ISPField[];
}

export interface ISPItemsResponse {
    value: ISPListItem[];
}
