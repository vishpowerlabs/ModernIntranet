import * as React from 'react';
import styles from './Faq.module.scss';
import { IFaqProps, IFaqItem } from './IFaqProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Spinner, SpinnerSize } from '@fluentui/react';
import FaqAccordion from './FaqAccordion';
import { ThemeService } from '../../../common/services/ThemeService';
import { EmptyState } from '../../../common/components/EmptyState/EmptyState';
import { WebPartHeader } from '../../../common/components/WebPartHeader/WebPartHeader';

const Faq: React.FC<IFaqProps> = (props) => {
    const [items, setItems] = React.useState<IFaqItem[]>([]);
    const [filteredItems, setFilteredItems] = React.useState<IFaqItem[]>([]);
    const [categories, setCategories] = React.useState<string[]>([]);
    const [loading, setLoading] = React.useState<boolean>(true);
    const [searchQuery, setSearchQuery] = React.useState<string>('');
    const [activeCategory, setActiveCategory] = React.useState<string>('All');
    const [expandedIds, setExpandedIds] = React.useState<Set<number>>(new Set());

    const fetchItems = async (): Promise<void> => {
        if (!props.listId) {
            setLoading(false);
            return;
        }

        setLoading(true);
        try {
            const selectFields = ['Id', props.questionColumn, props.answerColumn];
            if (props.categoryColumn) selectFields.push(props.categoryColumn);
            if (props.orderColumn) selectFields.push(props.orderColumn);

            const orderBy = props.orderColumn ? `&$orderby=${props.orderColumn} asc` : '';
            const url = `${props.siteUrl}/_api/web/lists(guid'${props.listId}')/items?$select=${selectFields.map(f => f?.trim()).filter(Boolean).join(',')}&$top=500${orderBy}`;

            const response: SPHttpClientResponse = await props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
            const data = await response.json();

            if (data.value) {
                const fetchedItems: IFaqItem[] = data.value.map((item: any) => ({
                    id: item.Id,
                    question: item[props.questionColumn],
                    answer: item[props.answerColumn],
                    category: props.categoryColumn ? item[props.categoryColumn] : undefined,
                    order: props.orderColumn ? item[props.orderColumn] : undefined
                }));

                setItems(fetchedItems);

                if (props.categoryColumn) {
                    const uniqueCats = Array.from(new Set(fetchedItems.map(i => i.category).filter(Boolean))) as string[];
                    setCategories([...uniqueCats].sort((a, b) => a.localeCompare(b)));
                }

                if (props.expandFirstItem && fetchedItems.length > 0) {
                    setExpandedIds(new Set([fetchedItems[0].id]));
                }
            }
        } catch (error) {
            console.error("Error fetching FAQ items:", error);
        } finally {
            setLoading(false);
        }
    };

    const filterData = (): void => {
        let filtered = items;

        if (activeCategory !== 'All') {
            filtered = filtered.filter(item => item.category === activeCategory);
        }

        if (searchQuery) {
            const q = searchQuery.toLowerCase();
            filtered = filtered.filter(item =>
                item.question?.toLowerCase().includes(q) ||
                item.answer?.toLowerCase().includes(q)
            );
        }

        setFilteredItems(filtered);
    };

    React.useEffect(() => {
        ThemeService.initialize(props.context);
        const load = async (): Promise<void> => {
            await fetchItems();
        };
        load().catch(err => console.error(err));
    }, [props.siteUrl, props.listId, props.questionColumn, props.answerColumn, props.categoryColumn, props.orderColumn]);

    React.useEffect(() => {
        const runFilter = (): void => {
            filterData();
        };
        runFilter();
    }, [searchQuery, activeCategory, items]);

    const handleSearch = (e: React.ChangeEvent<HTMLInputElement>): void => {
        setSearchQuery(e.target.value);
    };

    const clearFilters = (): void => {
        setSearchQuery('');
        setActiveCategory('All');
    };

    const toggleItem = (id: number): void => {
        setExpandedIds(prev => {
            const next = new Set(prev);
            if (next.has(id)) {
                next.delete(id);
            } else {
                if (!props.allowMultipleOpen) {
                    next.clear();
                }
                next.add(id);
            }
            return next;
        });
    };

    const renderHeader = (): JSX.Element => (
        <WebPartHeader
            title={props.title || ''}
            showTitle={!!props.showTitle}
            showBackgroundBar={!!props.showBackgroundBar}
            titleBarStyle={props.titleBarStyle || 'underline'}
        />
    );

    if (loading) {
        return <Spinner size={SpinnerSize.large} label="Loading FAQs..." />;
    }

    if (!props.listId || !props.questionColumn || !props.answerColumn) {
        return (
            <section className={styles.faq}>
                <EmptyState
                    icon="Questionnaire"
                    title="FAQ - Configuration Required"
                    message="Please complete the web part configuration to display FAQs."
                    description="You need to specify the Site URL, List ID, and map the required columns (Question, Answer) in the property pane."
                />
            </section>
        );
    }

    return (
        <section className={styles.faq}>
            {renderHeader()}

            <div className={styles.container}>
                {!props.showTitle && (
                    <div className={styles.header}>
                        <h1 className={styles.title}>{props.title}</h1>
                    </div>
                )}
                <div className={styles.headerInfo}>
                    <span className={styles.subtitle}>{filteredItems.length} questions</span>
                </div>

                {props.showSearch && (
                    <div className={styles.searchBox}>
                        <svg className={styles.searchIcon} viewBox="0 0 24 24">
                            <path d="M15.5 14h-.79l-.28-.27A6.471 6.471 0 0016 9.5 6.5 6.5 0 109.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z" />
                        </svg>
                        <input
                            type="text"
                            placeholder="Search questions…"
                            value={searchQuery}
                            onChange={handleSearch}
                            aria-label="Search frequently asked questions"
                        />
                        {searchQuery && (
                            <button className={styles.clearButton} onClick={() => setSearchQuery('')} title="Clear search">
                                ✕
                            </button>
                        )}
                    </div>
                )}

                {props.showCategoryFilter && categories.length > 0 && (
                    <div className={styles.pills} role="tablist">
                        <button
                            className={`${styles.pill} ${activeCategory === 'All' ? styles.active : ''}`}
                            onClick={() => setActiveCategory('All')}
                            role="tab"
                            aria-selected={activeCategory === 'All'}
                        >
                            All
                        </button>
                        {categories.map(cat => (
                            <button
                                key={cat}
                                className={`${styles.pill} ${activeCategory === cat ? styles.active : ''}`}
                                onClick={() => setActiveCategory(cat)}
                                role="tab"
                                aria-selected={activeCategory === cat}
                            >
                                {cat}
                            </button>
                        ))}
                    </div>
                )}

                {filteredItems.length > 0 ? (
                    <FaqAccordion
                        items={filteredItems}
                        expandedIds={expandedIds}
                        toggleItem={toggleItem}
                        searchQuery={searchQuery}
                    />
                ) : (
                    <div className={styles.noResults}>
                        No questions match your search.{' '}
                        <button
                            className={styles.clearLink}
                            onClick={clearFilters}
                            onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') clearFilters(); }}
                        >
                            Clear filters
                        </button>
                    </div>
                )}
            </div>
        </section>
    );
};

export default Faq;
