/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './ModernDocumentViewer.module.scss';
import { IModernDocumentViewerProps } from './IModernDocumentViewerProps';
import {
    SearchBox,
    MessageBar,
    MessageBarType,
    Spinner,
    SpinnerSize,
    Stack
} from '@fluentui/react';
import { SiteListService } from '../../../common/services/SiteListService';
import { SideNav } from './SideNav';
import { TopTabs } from './TopTabs';
import { DataTable } from './DataTable';
import { IDocument } from '../../../common/models';
import { EmptyState } from '../../../common/components/EmptyState/EmptyState';

export const ModernDocumentViewer: React.FunctionComponent<IModernDocumentViewerProps> = (props) => {
    const {
        context,
        siteUrl,
        listId,
        categoryField,
        subCategoryField,
        descriptionField,
        pinnedField,
        enableSubCategory,
        categoryDisplayType,
        webPartTitle,
        webPartTitleFontSize,
        webPartDescription,
        webPartDescriptionFontSize,
        pageSize
    } = props;

    const [items, setItems] = useState<IDocument[]>([]);
    const [loading, setLoading] = useState<boolean>(false);
    const [error, setError] = useState<string | undefined>(undefined);

    const [selectedCategory, setSelectedCategory] = useState<string>('All');
    const [selectedSubCategory, setSelectedSubCategory] = useState<string>('');
    const [searchText, setSearchText] = useState<string>('');

    const [categories, setCategories] = useState<string[]>([]);
    const [subCategories, setSubCategories] = useState<string[]>([]);
    const [filteredItems, setFilteredItems] = useState<IDocument[]>([]);

    const service = new SiteListService(context);

    useEffect(() => {
        const fetchDocuments = async (): Promise<void> => {
            if (!listId) return;

            setLoading(true);
            setError(undefined);

            try {
                const docs = await service.getDocuments(
                    siteUrl,
                    listId,
                    categoryField,
                    subCategoryField,
                    descriptionField,
                    pinnedField
                );

                setItems(docs);

                const docsCopy = [...docs];
                const uniqueCats = Array.from(new Set(docsCopy.map(d => d.Category).filter(Boolean)));
                uniqueCats.sort((a, b) => a.localeCompare(b));
                setCategories(['All', ...uniqueCats]);

            } catch (err) {
                setError((err as Error).message || "Failed to fetch documents");
            } finally {
                setLoading(false);
            }
        };

        fetchDocuments().catch(console.error);
    }, [siteUrl, listId, categoryField, subCategoryField, descriptionField, pinnedField]);

    useEffect(() => {
        let filtered = items;

        if (selectedCategory && selectedCategory !== 'All') {
            filtered = filtered.filter(i => i.Category === selectedCategory);
        }

        const subsCopy = Array.from(new Set(filtered.map(d => d.SubCategory).filter(Boolean)));
        subsCopy.sort((a, b) => a.localeCompare(b));
        setSubCategories(subsCopy);

        if (enableSubCategory && selectedSubCategory) {
            filtered = filtered.filter(i => i.SubCategory === selectedSubCategory);
        }

        if (searchText) {
            const lowerSearch = searchText.toLowerCase();
            filtered = filtered.filter(i =>
                i.Title?.toLowerCase().includes(lowerSearch) ||
                i.Description?.toLowerCase().includes(lowerSearch)
            );
        }

        setFilteredItems(filtered);
    }, [items, selectedCategory, selectedSubCategory, searchText]);

    if (!listId) {
        return (
            <section className={styles.documentListingV2}>
                <EmptyState icon="Document" message="Please configure a source library in the property pane." />
            </section>
        );
    }

    return (
        <div className={styles.documentListingV2}>
            {(webPartTitle || webPartDescription) && (
                <div className={styles.webpartHeader}>
                    <div className={styles.titleContainer}>
                        {webPartTitle && (
                            <h2 style={{ fontSize: webPartTitleFontSize }}>
                                {webPartTitle}
                            </h2>
                        )}
                        <div className={styles.backgroundBar} />
                    </div>
                    {webPartDescription && (
                        <div className={styles.headerDescription} style={{ fontSize: webPartDescriptionFontSize }}>
                            {webPartDescription}
                        </div>
                    )}
                </div>
            )}

            {loading && <Spinner size={SpinnerSize.large} label="Loading documents..." />}

            {!loading && !error && (
                <Stack horizontal wrap tokens={{ childrenGap: 20 }} styles={{ root: { paddingTop: 20, width: '100%' } }}>
                    {(enableSubCategory || categoryDisplayType === 'side') && (
                        <Stack.Item grow={1} styles={{ root: { minWidth: 200, maxWidth: 250 } }}>
                            <SideNav
                                categories={categories}
                                selectedCategory={selectedCategory}
                                onSelectCategory={(cat) => {
                                    setSelectedCategory(cat);
                                    setSelectedSubCategory('');
                                }}
                            />
                        </Stack.Item>
                    )}

                    <Stack.Item grow={3} styles={{ root: { minWidth: 300 } }}>
                        <Stack tokens={{ childrenGap: 10 }}>
                            <SearchBox
                                placeholder="Search by title or description"
                                onChange={(_, newValue) => setSearchText(newValue || '')}
                            />

                            {enableSubCategory ? (
                                <TopTabs
                                    subCategories={subCategories}
                                    selectedSubCategory={selectedSubCategory}
                                    onSelectSubCategory={setSelectedSubCategory}
                                />
                            ) : (
                                categoryDisplayType === 'top' && (
                                    <TopTabs
                                        subCategories={categories}
                                        selectedSubCategory={selectedCategory}
                                        onSelectSubCategory={setSelectedCategory}
                                    />
                                )
                            )}

                            <DataTable
                                items={filteredItems}
                                pageSize={pageSize}
                            />
                        </Stack>
                    </Stack.Item>
                </Stack>
            )}

            {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
        </div>
    );
};
