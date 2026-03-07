import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, SelectionMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { getFileTypeIconProps } from '@uifabric/file-type-icons';
import { IDocument } from '../../models/interfaces';

export interface IDataTableProps {
    items: IDocument[];
    onRequestAccess: (item: IDocument) => void;
    pageSize?: number;
    headerOpacity?: number;
    themePrimary?: string;
    headerTextColor?: string;
    showRequestAccess?: boolean;
}

import styles from '../DocumentListingV2.module.scss';


export const DataTable: React.FunctionComponent<IDataTableProps> = (props) => {
    const { items, onRequestAccess, pageSize = 10, headerOpacity, themePrimary, headerTextColor } = props;

    const [columns, setColumns] = React.useState<IColumn[]>([]);
    const [sortedItems, setSortedItems] = React.useState<IDocument[]>(items);
    const [currentPage, setCurrentPage] = React.useState<number>(1);

    // Use refs to access current state inside the stable callback
    const columnsRef = React.useRef<IColumn[]>([]);
    const sortedItemsRef = React.useRef<IDocument[]>(items);

    // Update refs on render
    columnsRef.current = columns;
    sortedItemsRef.current = sortedItems;

    // Stable Sorting Handler
    const _onColumnClick = React.useCallback((ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        const currentColumns = columnsRef.current;
        const currentItems = sortedItemsRef.current;

        const newColumns: IColumn[] = currentColumns.slice();
        const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];

        newColumns.forEach((newCol: IColumn) => {
            if (newCol === currColumn) {
                currColumn.isSortedDescending = !currColumn.isSortedDescending;
                currColumn.isSorted = true;
            } else {
                newCol.isSorted = false;
                newCol.isSortedDescending = true;
            }
        });

        const newSortedItems = [...currentItems].sort((a: IDocument, b: IDocument) => {
            // Priority: Pinned > Unpinned
            if (a.Pinned !== b.Pinned) {
                return a.Pinned ? -1 : 1;
            }

            const sortKey = currColumn.fieldName as keyof IDocument || 'Title';
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const aVal = (a as any)[sortKey];
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const bVal = (b as any)[sortKey];

            // Handle undefined/null safety
            const safeA = aVal || '';
            const safeB = bVal || '';

            if (safeA < safeB) return currColumn.isSortedDescending ? 1 : -1;
            if (safeA > safeB) return currColumn.isSortedDescending ? -1 : 1;
            return 0;
        });

        setColumns(newColumns);
        setSortedItems(newSortedItems);
        setCurrentPage(1); // Reset to first page on sort
    }, []); // Empty dependency as we use refs

    // Prepare Columns
    React.useEffect(() => {
        const cols: IColumn[] = [
            {
                key: 'pinned',
                name: '', // Empty header
                fieldName: 'Pinned',
                minWidth: 20,
                maxWidth: 20,
                isResizable: false,
                isIconOnly: true,
                onRender: (item: IDocument) => {
                    if (item.Pinned) {
                        return (
                            <Icon
                                iconName="Pin"
                                style={{
                                    color: themePrimary || '#0078d4',
                                    fontSize: 14,
                                    fontWeight: 'bold',
                                    display: 'flex',
                                    alignItems: 'center',
                                    height: '100%'
                                }}
                            />
                        );
                    }
                    return null;
                }
            },
            {
                key: 'fileType',
                name: 'Type',
                fieldName: 'Name', // for sorting context probably not strictly needed but good
                minWidth: 24,
                maxWidth: 24,
                isIconOnly: true,
                onRender: (item: IDocument) => {
                    const extension = item.Name.split('.').pop();
                    const iconProps = getFileTypeIconProps({ extension: extension, size: 24, imageFileType: 'png' });
                    return <Icon {...iconProps} />;
                }
            },
            {
                key: 'title',
                name: 'Title',
                fieldName: 'Title',
                minWidth: 150,
                maxWidth: 250,
                isRowHeader: true,
                isResizable: true,
                isSorted: true,
                isSortedDescending: false,
                onColumnClick: _onColumnClick,
                onRender: (item: IDocument) => {
                    return (
                        <span title={item.Title} style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', display: 'block' }}>
                            {item.Title}
                        </span>
                    );
                }
            },
            {
                key: 'description',
                name: 'Description',
                fieldName: 'Description',
                minWidth: 200,
                maxWidth: 350,
                isResizable: true,
                onColumnClick: _onColumnClick,
            },
            {
                key: 'modified',
                name: 'Modified',
                fieldName: 'Modified',
                minWidth: 100,
                maxWidth: 150,
                isResizable: true,
                onColumnClick: _onColumnClick,
                onRender: (item: IDocument) => {
                    return <span>{new Date(item.Modified).toLocaleDateString()}</span>;
                }
            }
        ];

        if (props.showRequestAccess) {
            cols.push({
                key: 'action',
                name: 'Action',
                minWidth: 50,
                maxWidth: 50,
                onRender: (item: IDocument) => {
                    return (
                        <TooltipHost content="Request Access">
                            <IconButton
                                iconProps={{ iconName: 'Mail' }}
                                onClick={() => onRequestAccess(item)}
                            />
                        </TooltipHost>
                    );
                }
            });
        }

        setColumns(cols);
    }, [_onColumnClick, props.showRequestAccess]); // Removed [] dependency to allow updates if handler changed (though it is stable now)

    // Initial load sort
    React.useEffect(() => {
        if (items.length > 0 && columns.length === 0) {
            // Initial sort logic if needed, or just let the default columns stand
        }
    }, [items, columns]);

    // Update sorted items when props change
    React.useEffect(() => {
        // Default sort: Pinned first
        const initSorted = [...items].sort((a, b) => {
            if (a.Pinned !== b.Pinned) {
                return a.Pinned ? -1 : 1;
            }
            return 0;
        });
        setSortedItems(initSorted);
        setCurrentPage(1);
    }, [items]);

    // Pagination Logic
    const pageCount = Math.ceil(sortedItems.length / pageSize);
    const pagedItems = sortedItems.slice((currentPage - 1) * pageSize, currentPage * pageSize);

    return (
        <div
            className={styles.dataTableWrapper}
            style={{
                '--headerOpacity': headerOpacity,
                '--headerTextColor': headerTextColor,
                ...(themePrimary ? { '--themePrimary': themePrimary } : {})
            } as React.CSSProperties}
        >
            <DetailsList
                items={pagedItems}
                columns={columns}
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
            />
            {pageCount > 1 && (
                <div style={{ display: 'flex', justifyContent: 'center', marginTop: 10, gap: 10 }}>
                    <IconButton
                        iconProps={{ iconName: 'ChevronLeft' }}
                        disabled={currentPage === 1}
                        onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
                    />
                    <span>Page {currentPage} of {pageCount}</span>
                    <IconButton
                        iconProps={{ iconName: 'ChevronRight' }}
                        disabled={currentPage === pageCount}
                        onClick={() => setCurrentPage(prev => Math.min(prev + 1, pageCount))}
                    />
                </div>
            )}
        </div>
    );
};
