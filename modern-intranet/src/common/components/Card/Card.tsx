import * as React from 'react';
import styles from './Card.module.scss';

export interface ICardProps {
    imageUrl?: string;
    onClick?: () => void;
    className?: string;
    children?: React.ReactNode;
}

export const Card: React.FC<ICardProps> = ({ imageUrl, onClick, className, children }) => {
    const rootClass = [
        styles.card,
        onClick ? styles.clickable : '',
        className || ''
    ].filter(Boolean).join(' ');

    const onKeyDown = (event: React.KeyboardEvent<HTMLButtonElement>): void => {
        if (onClick && (event.key === 'Enter' || event.key === ' ')) {
            event.preventDefault();
            onClick();
        }
    };

    const commonProps = {
        className: rootClass,
        onClick: onClick
    };

    if (onClick) {
        return (
            <button
                {...commonProps}
                onKeyDown={onKeyDown}
                type="button"
            >
                {imageUrl && (
                    <div
                        className={styles.imageArea}
                        style={{ backgroundImage: `url('${imageUrl}')` }}
                    />
                )}
                <div className={styles.bodyArea}>
                    {children}
                </div>
            </button>
        );
    }

    return (
        <div {...commonProps}>
            {imageUrl && (
                <div
                    className={styles.imageArea}
                    style={{ backgroundImage: `url('${imageUrl}')` }}
                />
            )}
            <div className={styles.bodyArea}>
                {children}
            </div>
        </div>
    );
};
