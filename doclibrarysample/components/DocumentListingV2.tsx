
import * as React from 'react';
import styles from './DocumentListingV2.module.scss';
import { IDocumentListingV2Props } from './IDocumentListingV2Props';

import {
  SearchBox,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
  Stack
} from '@fluentui/react';
import { SPService } from '../services/SPService';
import { SideNav } from './Navigation/SideNav';
import { TopTabs } from './Navigation/TopTabs';
import { DataTable } from './Content/DataTable';
import { IDocument } from '../models/interfaces';

export const DocumentListingV2: React.FunctionComponent<IDocumentListingV2Props> = (props) => {
  const {
    context,
    sourceLibraryId,
    categoryField,
    subCategoryField,
    descriptionField,
    webPartTitle,
    webPartTitleFontSize,
    webPartDescription,
    webPartDescriptionFontSize,
    headerOpacity,
    themePrimary,
    headerTextColor,
    pinnedField,
    showRequestAccess,
    siteUrl,
    requestSiteUrl
  } = props;

  const [items, setItems] = React.useState<IDocument[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [message, setMessage] = React.useState<{ text: string, type: MessageBarType } | undefined>(undefined);

  // Filters
  const [selectedCategory, setSelectedCategory] = React.useState<string>('All');
  const [selectedSubCategory, setSelectedSubCategory] = React.useState<string>('');
  const [searchText, setSearchText] = React.useState<string>('');

  // Derived Lists
  const [categories, setCategories] = React.useState<string[]>([]);
  const [subCategories, setSubCategories] = React.useState<string[]>([]);
  const [filteredItems, setFilteredItems] = React.useState<IDocument[]>([]);

  // Initialize Service
  React.useEffect(() => {
    if (context) {
      SPService.Instance.setup(context);
    }
  }, [context]);

  // Fetch Data
  React.useEffect(() => {
    if (sourceLibraryId) {
      setLoading(true);
      SPService.Instance.getDocuments(sourceLibraryId, categoryField, subCategoryField, descriptionField, pinnedField, siteUrl)
        .then((docs: IDocument[]) => {
          setItems(docs);

          // Extract Categories
          const uniqueCats = Array.from(new Set(docs.map(d => d.Category).filter(Boolean)));
          setCategories(uniqueCats.sort());
          if (uniqueCats.length > 0) setSelectedCategory(uniqueCats[0]);

          setLoading(false);
        })
        .catch((err: any) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const errorMessage = err.message || (typeof err === 'string' ? err : 'Unknown error');
          setError(errorMessage);
          setLoading(false);
        });
    }
  }, [sourceLibraryId, categoryField, subCategoryField, descriptionField, pinnedField, siteUrl]);

  // Filter Logic
  React.useEffect(() => {
    let filtered = items;

    // 1. Category
    if (selectedCategory && selectedCategory !== 'All') {
      filtered = filtered.filter(i => i.Category === selectedCategory);
    }

    // Update SubCategories based on Category (before subcat filter)
    const uniqueSubCats = Array.from(new Set(filtered.map(d => d.SubCategory).filter(Boolean)));
    setSubCategories(uniqueSubCats.sort());

    // 2. SubCategory
    if (selectedSubCategory) {
      filtered = filtered.filter(i => i.SubCategory === selectedSubCategory);
    }

    // 3. Search
    if (searchText) {
      const lowerSearch = searchText.toLowerCase();
      filtered = filtered.filter(i =>
        i.Title?.toLowerCase().includes(lowerSearch) ||
        i.Description?.toLowerCase().includes(lowerSearch)
      );
    }

    setFilteredItems(filtered);

  }, [items, selectedCategory, selectedSubCategory, searchText]);

  // Handlers
  /* State for Dialog */
  const [showReminderDialog, setShowReminderDialog] = React.useState<boolean>(false);
  const [duplicateRequestId, setDuplicateRequestId] = React.useState<number | undefined>(undefined);

  /* Dialog Logic */
  const handleRequestAccess = async (item: IDocument): Promise<void> => {
    const requestListId = props.requestListId;
    if (!requestListId) {
      setMessage({ text: 'Request List not configured.', type: MessageBarType.error });
      return;
    }

    try {
      const emailField = props.requestEmailField || 'Email';
      const fileIdField = props.requestFileIdField || 'FileID';
      const requestIdField = props.requestRequestIdField || 'RequestID';
      const dateField = props.requestDateField || 'RequestDate';

      const result = await SPService.Instance.logAccessRequest(
        requestListId,
        context.pageContext.user.email,
        item.Id.toString(),
        emailField,
        fileIdField,
        requestIdField,
        dateField,
        requestSiteUrl
      );

      if (result.status === 'Exists') {
        setDuplicateRequestId(result.itemId);
        setShowReminderDialog(true);
      } else {
        setMessage({ text: `Access requested for ${item.Title}`, type: MessageBarType.success });
        setTimeout(() => setMessage(undefined), 5000); // Clear message
      }

    } catch {
      setMessage({ text: 'Failed to request access.', type: MessageBarType.error });
    }
  };

  const handleSendReminder = async (): Promise<void> => {
    if (duplicateRequestId && props.requestListId && props.requestReminderField) {
      try {
        await SPService.Instance.setRequestReminder(
          props.requestListId,
          duplicateRequestId,
          props.requestReminderField,
          requestSiteUrl
        );
        setMessage({ text: props.reminderSentMessage || 'Reminder sent!', type: MessageBarType.success });
      } catch {
        setMessage({ text: 'Failed to send reminder.', type: MessageBarType.error });
      }
      setShowReminderDialog(false);
      setDuplicateRequestId(undefined);
      setTimeout(() => setMessage(undefined), 5000);
    } else {
      setShowReminderDialog(false); // Close if config missing
    }
  };

  const closeDialog = (): void => {
    setShowReminderDialog(false);
    setDuplicateRequestId(undefined);
  };

  const dialogContentProps = {
    type: DialogType.normal,
    title: 'Access Already Requested',
    closeButtonAriaLabel: 'Close',
    subText: props.alreadyRequestedMessage || 'You have already requested access. Would you like to send the information again?'
  };

  const dialogModalProps = {
    isBlocking: false,
    styles: { main: { maxWidth: 450 } }
  };

  if (!sourceLibraryId) {
    return <MessageBar messageBarType={MessageBarType.warning}>Please configure a source library in the property pane.</MessageBar>;
  }

  return (
    <div className={styles.documentListingV2}>
      {/* Title and Description Row */}
      {(webPartTitle || webPartDescription) && (
        <div
          className={styles.headerBar}
          style={{
            '--headerOpacity': headerOpacity,
            '--headerTextColor': headerTextColor || 'var(--bodyText)',
            ...(themePrimary ? { '--themePrimary': themePrimary } : {})
          } as React.CSSProperties}
        >
          {webPartTitle && (
            <div
              className={styles.headerTitle}
              style={{
                fontSize: webPartTitleFontSize || '24px',
                color: 'var(--headerTextColor)',
                marginBottom: webPartDescription ? '8px' : '0'
              }}
            >
              {webPartTitle}
            </div>
          )}
          {webPartDescription && (
            <div
              className={styles.headerDescription}
              style={{
                fontSize: webPartDescriptionFontSize || '14px',
                color: 'var(--headerTextColor)'
              }}
            >
              {webPartDescription}
            </div>
          )}
        </div>
      )}

      {message && (
        <MessageBar
          messageBarType={message.type}
          onDismiss={() => setMessage(undefined)}
          dismissButtonAriaLabel="Close"
        >
          {message.text}
        </MessageBar>
      )}

      {loading && <Spinner size={SpinnerSize.large} label="Loading documents..." />}

      {!loading && !error && (
        <Stack horizontal wrap tokens={{ childrenGap: 20 }} styles={{ root: { paddingTop: 20, width: '100%' } }}>
          <Stack.Item
            grow={1}
            styles={{
              root: {
                minWidth: 200,
                maxWidth: 300,
                '@media(max-width: 640px)': { width: '100%', maxWidth: '100%' }
              }
            }}
          >
            <SideNav
              categories={categories}
              selectedCategory={selectedCategory}
              onSelectCategory={(cat) => {
                setSelectedCategory(cat);
                setSelectedSubCategory(''); // Reset subcat
              }}
            />
          </Stack.Item>

          <Stack.Item grow={3} styles={{ root: { minWidth: 300 } }}>
            <Stack tokens={{ childrenGap: 10 }}>
              <SearchBox
                placeholder="Search by title or description"
                onChange={(_, newValue) => setSearchText(newValue || '')}
                styles={{ root: { width: '100%' } }}
              />

              <TopTabs
                subCategories={subCategories}
                selectedSubCategory={selectedSubCategory}
                onSelectSubCategory={setSelectedSubCategory}
              />

              <DataTable
                items={filteredItems}
                onRequestAccess={handleRequestAccess}
                pageSize={props.pageSize}
                headerOpacity={headerOpacity}
                themePrimary={themePrimary}
                headerTextColor={headerTextColor}
                showRequestAccess={showRequestAccess}
              />
            </Stack>
          </Stack.Item>
        </Stack>
      )}

      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}

      <Dialog
        hidden={!showReminderDialog}
        onDismiss={closeDialog}
        dialogContentProps={dialogContentProps}
        modalProps={dialogModalProps}
      >
        <DialogFooter>
          <PrimaryButton onClick={handleSendReminder} text="Send Info Again" />
          <DefaultButton onClick={closeDialog} text="Cancel" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};
