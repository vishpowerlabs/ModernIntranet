export interface IDocument {
    Id: number;
    Title: string;
    Description: string;
    FileRef: string; // URL
    Pinned?: boolean;
    Modified: string;
    Category: string; // from Dropdown
    SubCategory: string; // from Dropdown
    UniqueId: string; // For key
    Name: string; // File name
    ServerRelativeUrl: string;
}

export interface IAccessRequest {
    Title: string; // User Email + File ID
    Email: string;
    FileID: string;
    Count: number;
}
