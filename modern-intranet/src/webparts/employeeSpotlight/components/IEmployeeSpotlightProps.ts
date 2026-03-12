/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IEmployeeSpotlightProps {
    siteUrl: string;
    listId: string;

    // Column Mappings
    nameColumn: string;
    photoColumn: string;
    jobTitleColumn: string;
    departmentColumn: string;
    emailColumn: string;
    spotlightColumn: string;
    spotlightTextColumn: string;

    // Graph Source Properties
    source?: 'spList' | 'graph';
    selectedUsers?: any[];
    commonDescription?: string;

    // Display Settings
    maxItems: number;
    autoRotateInterval: number;

    // Header Settings
    showTitle: boolean;
    title: string;
    webPartTitleFontSize: string;
    showBackgroundBar: boolean;
    titleBarStyle: 'solid' | 'underline';
    layoutMode?: 'standard' | 'compact';

    context: WebPartContext;
}
