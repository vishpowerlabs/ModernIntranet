import * as React from 'react';
import styles from './Avatar.module.scss';
import { stringToColor } from '../../utils/colorUtils';

export interface IAvatarProps {
    name: string;
    size?: number;
    color?: string;
}

export const Avatar: React.FC<IAvatarProps> = ({ name, size = 36, color }) => {
    const getInitials = (fullName: string): string => {
        if (!fullName) return '';
        const parts = fullName.split(' ').filter(p => p.length > 0);
        if (parts.length === 0) return '';
        if (parts.length === 1) return parts[0].substring(0, 1).toUpperCase();
        return (parts[0].substring(0, 1) + parts[parts.length - 1].substring(0, 1)).toUpperCase();
    };

    const bgColor = color || stringToColor(name);
    const fontSize = Math.max(10, Math.floor(size / 2.5));

    return (
        <div
            className={styles.avatar}
            style={{
                width: size,
                height: size,
                backgroundColor: bgColor,
                fontSize: `${fontSize}px`
            }}
            title={name}
        >
            {getInitials(name)}
        </div>
    );
};
