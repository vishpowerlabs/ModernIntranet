/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './EmptyState.module.scss';

export interface IEmptyStateProps {
    message: string;
    icon?: string;
}

export const EmptyState: React.FC<IEmptyStateProps> = ({ message, icon }) => {
    return (
        <div className={styles.emptyState}>
            {icon && <Icon iconName={icon} className={styles.icon} />}
            <div className={styles.message}>{message}</div>
        </div>
    );
};
