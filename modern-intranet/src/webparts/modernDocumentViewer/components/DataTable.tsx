/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, SelectionMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import { IDocument } from '../../../common/models';
import styles from './ModernDocumentViewer.module.scss';

export interface IDataTableProps {
    items: IDocument[];
    pageSize?: number;
}

const renderPinned = (item: IDocument): JSX.Element | null => {
    if (item.Pinned) {
        return (
            <Icon
                iconName="Pin"
                style={{
                    color: 'var(--themePrimary)',
                    fontSize: 14,
                    fontWeight: 600,
                    display: 'flex',
                    alignItems: 'center',
                    height: '100%'
                }}
            />
        );
    }
    return null;
};

const renderFileType = (): JSX.Element => (
    <Icon
        iconName="Document"
        style={{ color: 'var(--themePrimary)', fontSize: 16 }}
    />
);

const renderTitle = (item: IDocument): JSX.Element => (
    <a
        href={`${item.EncodedAbsUrl}?web=1`}
        target="_blank"
        rel="noopener noreferrer"
        title={item.Title}
        style={{
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap',
            display: 'block',
            color: 'var(--bodyText)',
            textDecoration: 'none'
        }}
    >
        {item.Title}
    </a>
);

const renderModified = (item: IDocument): JSX.Element => (
    <span>{new Date(item.Modified).toLocaleDateString()}</span>
);

export const DataTable: React.FunctionComponent<IDataTableProps> = (props) => {
    const { items, pageSize = 10 } = props;

    const [columns, setColumns] = React.useState<IColumn[]>([]);
    const [sortedItems, setSortedItems] = React.useState<IDocument[]>(items);
    const [currentPage, setCurrentPage] = React.useState<number>(1);

    const columnsRef = React.useRef<IColumn[]>([]);
    const sortedItemsRef = React.useRef<IDocument[]>(items);

    columnsRef.current = columns;
    sortedItemsRef.current = sortedItems;

    const _onColumnClick = React.useCallback((ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        const currentColumns = columnsRef.current;
        const currentItems = sortedItemsRef.current;

        const newColumns: IColumn[] = currentColumns.slice();
        const currColumn: IColumn | undefined = newColumns.find(currCol => column.key === currCol.key);

        if (!currColumn) return;

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
            if (a.Pinned !== b.Pinned) {
                return a.Pinned ? -1 : 1;
            }

            const sortKey = currColumn.fieldName as keyof IDocument || 'Title';
            const aVal = (a as any)[sortKey];
            const bVal = (b as any)[sortKey];

            const safeA = aVal || '';
            const safeB = bVal || '';

            if (safeA < safeB) return currColumn.isSortedDescending ? 1 : -1;
            if (safeA > safeB) return currColumn.isSortedDescending ? -1 : 1;
            return 0;
        });

        setColumns(newColumns);
        setSortedItems(newSortedItems);
        setCurrentPage(1);
    }, []);

    React.useEffect(() => {
        const cols: IColumn[] = [
            {
                key: 'pinned',
                name: '',
                fieldName: 'Pinned',
                minWidth: 20,
                maxWidth: 20,
                isResizable: false,
                isIconOnly: true,
                onRender: renderPinned
            },
            {
                key: 'fileType',
                name: 'Type',
                fieldName: 'Name',
                minWidth: 24,
                maxWidth: 24,
                isIconOnly: true,
                onRender: renderFileType
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
                onRender: renderTitle
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
                onRender: renderModified
            }
        ];

        setColumns(cols);
    }, [_onColumnClick]);

    React.useEffect(() => {
        const initSorted = [...items].sort((a, b) => {
            if (a.Pinned !== b.Pinned) {
                return a.Pinned ? -1 : 1;
            }
            return 0;
        });
        setSortedItems(initSorted);
        setCurrentPage(1);
    }, [items]);

    const pageCount = Math.ceil(sortedItems.length / pageSize);
    const pagedItems = sortedItems.slice((currentPage - 1) * pageSize, currentPage * pageSize);

    return (
        <div className={styles.dataTableWrapper}>
            <DetailsList
                items={pagedItems}
                columns={columns}
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={(item: IDocument) => {
                    window.open(`${item.EncodedAbsUrl}?web=1`, '_blank');
                }}
            />
            {pageCount > 1 && (
                <div className={styles.pagination}>
                    <IconButton
                        iconProps={{
                            iconName: 'ChevronLeft',
                            style: { color: 'var(--themePrimary)' }
                        }}
                        disabled={currentPage === 1}
                        onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
                    />
                    <span>Page {currentPage} of {pageCount}</span>
                    <IconButton
                        iconProps={{
                            iconName: 'ChevronRight',
                            style: { color: 'var(--themePrimary)' }
                        }}
                        disabled={currentPage === pageCount}
                        onClick={() => setCurrentPage(prev => Math.min(prev + 1, pageCount))}
                    />
                </div>
            )}
        </div>
    );
};
