import { useState, useEffect } from 'react';
import { IColumnInfo } from '../models';
import { SiteListService } from '../services/SiteListService';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export function useColumns(context: WebPartContext, siteUrl: string, listId: string, typeFilter?: string): IColumnInfo[] {
    const [columns, setColumns] = useState<IColumnInfo[]>([]);

    useEffect(() => {
        if (!context || !siteUrl || !listId) {
            setColumns([]);
            return;
        }

        let isMounted = true;
        const service = new SiteListService(context);

        service.getColumns(siteUrl, listId, typeFilter).then(fetchedColumns => {
            if (isMounted) {
                setColumns(fetchedColumns || []);
            }
        }).catch(error => {
            console.error('useColumns error:', error);
        });

        return () => {
            isMounted = false;
        };
    }, [context, siteUrl, listId, typeFilter]);

    return columns;
}
