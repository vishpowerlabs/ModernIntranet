/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import styles from './BannerSlider.module.scss';

export interface ISlideProps {
    title: string;
    description?: string;
    imageUrl: string;
    buttonText?: string;
    pageLink?: string;
    showCta: boolean;
    isActive: boolean;
}

export const Slide: React.FC<ISlideProps> = ({
    title,
    description,
    imageUrl,
    buttonText,
    pageLink,
    showCta,
    isActive
}) => {
    return (
        <div className={styles.bannerSlide}>
            <div
                className={styles.imageSection}
                style={{ backgroundImage: `url('${imageUrl}')` }}
            >
                {showCta && buttonText && pageLink && (
                    <a href={pageLink} className={styles.ctaButton} target="_blank" rel="noopener noreferrer">
                        {buttonText}
                    </a>
                )}
            </div>
            <div className={styles.textSection}>
                {title && <h2 className={styles.slideTitle}>{title}</h2>}
                {description && <p className={styles.slideDescription}>{description}</p>}
            </div>
        </div>
    );
};
