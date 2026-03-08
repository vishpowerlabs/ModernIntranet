/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './EmptyState.module.scss';

export interface IEmptyStateProps {
    title?: string;
    message: string;
    description?: string;
    icon?: string;
}

export const EmptyState: React.FC<IEmptyStateProps> = ({ title, message, description, icon }) => {
    return (
        <div className={styles.emptyState}>
            {icon && <Icon iconName={icon} className={styles.icon} />}
            {title && <div className={styles.title}>{title}</div>}
            <div className={styles.message}>{message}</div>
            {description && <div className={styles.descriptionText}>{description}</div>}
        </div>
    );
};
