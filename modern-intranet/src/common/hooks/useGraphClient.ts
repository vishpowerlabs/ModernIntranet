import { useState, useEffect } from 'react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export function useGraphClient(context: WebPartContext): MSGraphClientV3 | null {
    const [client, setClient] = useState<MSGraphClientV3 | null>(null);

    useEffect(() => {
        if (!context?.msGraphClientFactory) return;

        let isMounted = true;
        context.msGraphClientFactory.getClient('3')
            .then((c: MSGraphClientV3) => {
                if (isMounted) {
                    setClient(c);
                }
            })
            .catch((error: Error) => {
                console.error('Error getting Graph client', error);
            });

        return () => {
            isMounted = false;
        };
    }, [context]);

    return client;
}
