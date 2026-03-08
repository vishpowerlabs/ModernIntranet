/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import styles from './EmployeeSpotlight.module.scss';

export interface ISpotlightItem {
    Id: number;
    Name: string;
    PhotoUrl?: string;
    JobTitle: string;
    Department: string;
    SpotlightText: string;
    Email: string;
}

interface ISpotlightCardProps {
    item: ISpotlightItem;
}

export const SpotlightCard: React.FC<ISpotlightCardProps> = ({ item }) => {

    // Fallback initials generator
    const getInitials = (name: string): string => {
        if (!name) return '??';
        const parts = name.split(' ');
        return parts.map(p => p[0]).join('').substring(0, 2).toUpperCase();
    };

    // Consistent hashing color for avatar background
    const getAvatarColor = (name: string): string => {
        const colors = [
            '#0078d4', '#038387', '#8764b8', '#c239b3',
            '#ca5010', '#498205', '#005b70', '#7160e8',
            '#d13438', '#4f6bed', '#986f0b', '#3b3a39'
        ];
        let hash = 0;
        for (let i = 0; i < name.length; i++) {
            hash = name.charCodeAt(i) + ((hash << 5) - hash);
        }
        return colors[Math.abs(hash) % colors.length];
    };

    return (
        <div className={styles.wpSpotlightCard}>
            <div className={styles.wpSpotlightPhoto}>
                {item.PhotoUrl ? (
                    <img src={item.PhotoUrl} alt={item.Name} />
                ) : (
                    <div
                        className={styles.wpSpotlightAvatarLg}
                        style={{ backgroundColor: getAvatarColor(item.Name) }}
                    >
                        {getInitials(item.Name)}
                    </div>
                )}
            </div>

            <div className={styles.wpSpotlightBody}>
                <div className={styles.wpSpotlightBadge}>
                    ⭐ Employee Spotlight
                </div>

                <div className={styles.wpSpotlightName}>
                    {item.Name}
                </div>

                <div className={styles.wpSpotlightRole}>
                    {item.JobTitle} {item.Department ? `· ${item.Department}` : ''}
                </div>

                <div className={styles.wpSpotlightText}>
                    {item.SpotlightText}
                </div>

                {item.Email && (
                    <a className={styles.wpSpotlightEmail} href={`mailto:${item.Email}`}>
                        ✉ {item.Email}
                    </a>
                )}
            </div>
        </div>
    );
};
