import { useState, useEffect } from 'react';
import { IListInfo } from '../models';
import { SiteListService } from '../services/SiteListService';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export function useLists(context: WebPartContext, siteUrl: string): IListInfo[] {
    const [lists, setLists] = useState<IListInfo[]>([]);

    useEffect(() => {
        if (!context || !siteUrl) {
            setLists([]);
            return;
        }

        let isMounted = true;
        const service = new SiteListService(context);

        service.getLists(siteUrl).then(fetchedLists => {
            if (isMounted) {
                setLists(fetchedLists || []);
            }
        }).catch(error => {
            console.error('useLists error:', error);
        });

        return () => {
            isMounted = false;
        };
    }, [context, siteUrl]);

    return lists;
}
