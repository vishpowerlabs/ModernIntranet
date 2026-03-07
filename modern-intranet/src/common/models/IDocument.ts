export interface IDocument {
    Id: number;
    Title: string;
    Name: string;
    Description: string;
    FileRef: string;
    Pinned: boolean;
    Modified: string;
    Created: string;
    Category: string;
    SubCategory: string;
    Author: string;
    Editor: string;
    UniqueId: string;
    ServerRelativeUrl: string;
    EncodedAbsUrl: string;
}
