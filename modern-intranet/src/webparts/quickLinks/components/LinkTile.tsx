/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './QuickLinks.module.scss';

export interface ILinkTileProps {
    title: string;
    url: string;
    icon?: string;
    openInNewTab: boolean;
}

export const LinkTile: React.FC<ILinkTileProps> = ({ title, url, icon, openInNewTab }) => {
    // If icon is provided and looks like it's alphanumeric/Fluent UI icon name, render Icon
    // Otherwise render it as text (emoji fallback)
    const isFluentIcon = icon && /^[a-zA-Z0-9-]+$/.test(icon);

    let iconElement: JSX.Element;
    if (icon) {
        if (isFluentIcon) {
            iconElement = <Icon iconName={icon} className={styles.fluentIcon} />;
        } else {
            iconElement = <span className={styles.emojiIcon}>{icon}</span>;
        }
    } else {
        iconElement = <Icon iconName="Link" className={styles.fluentIcon} />;
    }

    return (
        <a
            className={styles.linkTile}
            href={url || '#'}
            target={openInNewTab ? '_blank' : '_self'}
            rel={openInNewTab ? 'noopener noreferrer' : undefined}
            title={title}
        >
            <div className={styles.iconContainer}>
                {iconElement}
            </div>
            <div className={styles.label}>{title}</div>
        </a>
    );
};
