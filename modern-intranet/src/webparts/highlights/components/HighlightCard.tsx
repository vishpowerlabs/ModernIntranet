/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { Card } from '../../../common/components/Card/Card';
import styles from './Highlights.module.scss';

export interface IHighlightCardProps {
    title: string;
    description?: string;
    imageUrl?: string;
    linkUrl?: string;
}

export const HighlightCard: React.FC<IHighlightCardProps> = ({ title, description, imageUrl, linkUrl }) => {
    const onCardClick = (): void => {
        if (linkUrl) {
            window.open(linkUrl, '_blank', 'noopener noreferrer');
        }
    };

    return (
        <Card
            imageUrl={imageUrl} // Although Card has imageUrl, Highlights requirements specify 
            // banner image at top. Card.tsx uses imageUrl as background-image 
            // for styles.imageArea. Let's see if we can use it or customize.
            onClick={linkUrl ? onCardClick : undefined}
            className={styles.highlightCard}
        >
            <div className={styles.contentBody}>
                <a
                    href={linkUrl}
                    className={styles.title}
                    target="_blank"
                    rel="noopener noreferrer"
                    onClick={(e) => e.stopPropagation()} // Prevent double trigger if Card handles click
                >
                    {title}
                </a>
                {description && <p className={styles.description}>{description}</p>}
            </div>
        </Card>
    );
};
