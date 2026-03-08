import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFaqItem {
    id: number;
    question: string;
    answer: string;
    category?: string;
    order?: number;
}

export interface IFaqProps {
    context: WebPartContext;
    siteUrl: string;
    listId: string;

    questionColumn: string;
    answerColumn: string;
    categoryColumn?: string;
    orderColumn?: string;

    showTitle: boolean;
    title: string;
    showBackgroundBar: boolean;
    titleBarStyle: 'solid' | 'underline';

    showSearch: boolean;
    showCategoryFilter: boolean;
    allowMultipleOpen: boolean;
    expandFirstItem: boolean;
}
