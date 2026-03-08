/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IEventsProps {
  siteUrl: string;
  listId: string;
  titleColumn: string;
  dateColumn: string;
  activeColumn: string;
  imageColumn?: string;
  linkColumn?: string;
  locationColumn?: string;
  pinnedColumn?: string;
  maxItems: number;
  itemsPerRow: number;
  showViewAll: boolean;
  viewAllUrl?: string;
  title: string;
  showTitle: boolean;
  showBackgroundBar: boolean;
  titleBarStyle: 'solid' | 'underline';
  siteId: string;
  webId: string;
  context: WebPartContext;
}
