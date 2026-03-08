import * as React from 'react';
import styles from './NewJoiners.module.scss';

export interface INewJoiner {
    name: string;
    photoUrl: string;
    jobTitle: string;
    department: string;
    introText: string;
    email: string;
}

interface INewJoinerCardProps {
    joiner: INewJoiner;
    layout: string;
}

export const NewJoinerCard: React.FC<INewJoinerCardProps> = ({ joiner, layout }) => {

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
            hash = (name.codePointAt(i) || 0) + ((hash << 5) - hash);
        }
        return colors[Math.abs(hash) % colors.length];
    };

    // If layout is strip, we use horizontal card style
    const isHorizontal = layout === 'strip';
    const cardClass = `${styles.card} ${isHorizontal ? styles.horizontal : styles.vertical}`;

    return (
        <div className={cardClass}>
            <div className={styles.photoArea}>
                {joiner.photoUrl ? (
                    <img src={joiner.photoUrl} alt={joiner.name} />
                ) : (
                    <div
                        className={styles.avatarPlaceholder}
                        style={{ backgroundColor: getAvatarColor(joiner.name) }}
                    >
                        {getInitials(joiner.name)}
                    </div>
                )}
            </div>

            <div className={styles.cardBody}>
                <div className={styles.badge}>
                    👋 New Joiner
                </div>

                <div className={styles.name}>{joiner.name}</div>

                <div className={styles.role}>
                    {joiner.jobTitle} {joiner.department ? `· ${joiner.department}` : ''}
                </div>

                {joiner.introText && (
                    <div className={styles.intro} title={joiner.introText}>
                        {joiner.introText}
                    </div>
                )}
            </div>
        </div>
    );
};
