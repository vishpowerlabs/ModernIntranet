import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDocumentListingV2Props {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  // Custom Properties
  context: WebPartContext;
  sourceLibraryId: string;
  requestListId: string;
  siteUrl?: string;
  requestSiteUrl?: string;

  // Field Mappings (Internal Names)
  categoryField: string;
  subCategoryField: string;
  descriptionField: string;
  pageSize: number;
  requestEmailField: string;
  requestFileIdField: string;
  requestRequestIdField: string;
  requestDateField: string;
  requestReminderField: string;
  alreadyRequestedMessage: string;
  webPartTitle: string;
  webPartTitleFontSize: string;
  webPartDescription: string;
  webPartDescriptionFontSize: string;
  reminderSentMessage: string;
  headerOpacity: number;
  themePrimary?: string;
  headerTextColor?: string;
  pinnedField?: string;
  showRequestAccess?: boolean;
}
