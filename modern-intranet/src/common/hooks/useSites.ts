import { useState, useEffect } from 'react';
import { ISite } from '../models';
import { SiteListService } from '../services/SiteListService';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export function useSites(context: WebPartContext): ISite[] {
    const [sites, setSites] = useState<ISite[]>([]);

    useEffect(() => {
        if (!context) return;

        let isMounted = true;
        const service = new SiteListService(context);

        service.getSites().then(fetchedSites => {
            if (isMounted) {
                setSites(fetchedSites || []);
            }
        }).catch(error => {
            console.error('useSites error:', error);
        });

        return () => {
            isMounted = false;
        };
    }, [context]);

    return sites;
}
