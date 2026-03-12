import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IManualJoiner {
    user: any;
    introText: string;
}

export interface INewJoinersProps {
    siteUrl: string;
    listId: string;
    nameColumn: string;
    photoColumn: string;
    jobTitleColumn: string;
    departmentColumn: string;
    emailColumn: string;
    newJoinerColumn: string;
    newJoinerTextColumn: string;
    maxItems: number;
    layout: 'list' | 'grid' | 'strip';
    layoutMode: 'standard' | 'compact';
    source: 'spList' | 'graph';
    manualJoiners: any[];
    commonIntro: string;
    autoRotateInterval?: number;
    showTitle: boolean;
    title: string;
    showBackgroundBar: boolean;
    titleBarStyle: 'solid' | 'underline';
    context: WebPartContext;
}
