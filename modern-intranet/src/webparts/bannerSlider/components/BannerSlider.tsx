/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import styles from './BannerSlider.module.scss';
import { IBannerSliderProps } from './IBannerSliderProps';
import { Slide, ISlideProps } from './Slide';
import { EmptyState } from '../../../common/components/EmptyState/EmptyState';
import { SPHttpClient } from '@microsoft/sp-http';

interface ISharePointImageMetadata {
    fileName?: string;
    serverRelativeUrl?: string;
    serverUrl?: string;
    Url?: string;
}

interface ISharePointRow {
    ID?: string;
    FileDirRef?: string;
    [key: string]: string | number | boolean | undefined | null; // Improved over 'any'
}

const getPreviewUrl = (fileName: string, siteUrl: string, siteId: string, webId: string): string => {
    // Simplified regex to reduce SonarQube complexity
    const match = /([a-f\d-]{32,36})/i.exec(fileName);
    if (!match) return '';

    let guid = match[1];
    if (guid.length === 32) {
        guid = `${guid.slice(0, 8)}-${guid.slice(8, 12)}-${guid.slice(12, 16)}-${guid.slice(16, 20)}-${guid.slice(20)}`;
    }
    return `${siteUrl}/_layouts/15/getpreview.ashx?guidSite=${siteId}&guidWeb=${webId}&guidFile=${guid}&clientType=image`;
};

const resolveImageObject = (rawValue: string | object): ISharePointImageMetadata | null => {
    if (typeof rawValue === 'string' && rawValue.startsWith('{')) {
        try { return JSON.parse(rawValue); } catch { return null; }
    }
    return typeof rawValue === 'object' ? rawValue : null;
};

/**
 * Helper to extract image URL from various SharePoint Image column formats
 */
const getImageUrl = (rowItem: ISharePointRow, imageColumn: string, siteUrl: string, siteId: string, webId: string): string => {
    const rawValue = rowItem[imageColumn];
    if (!rawValue) return '';

    const imageObj = resolveImageObject(rawValue as string | object);
    if (!imageObj) return typeof rawValue === 'string' ? rawValue : '';

    const url = imageObj.serverRelativeUrl || imageObj.serverUrl || imageObj.Url;
    if (url) {
        return (!url.startsWith('http') && url.startsWith('/'))
            ? `${new URL(siteUrl).origin}${url}`
            : url;
    }

    if (imageObj.fileName && rowItem.FileDirRef && rowItem.ID) {
        return `${new URL(siteUrl).origin}${rowItem.FileDirRef}/Attachments/${rowItem.ID}/${imageObj.fileName}`;
    }

    return imageObj.fileName ? getPreviewUrl(imageObj.fileName, siteUrl, siteId, webId) : '';
};

export const BannerSlider: React.FC<IBannerSliderProps> = (props) => {
    const [slides, setSlides] = useState<ISlideProps[]>([]);
    const [currentIndex, setCurrentIndex] = useState(0);
    const [isLoading, setIsLoading] = useState(true);
    const [error, setError] = useState<string | null>(null);
    const timerRef = useRef<any>(null);

    const nextSlide = React.useCallback(() => {
        setCurrentIndex((prev) => (prev + 1) % (slides.length || 1));
    }, [slides.length]);

    const prevSlide = React.useCallback(() => {
        setCurrentIndex((prev) => (prev - 1 + slides.length) % (slides.length || 1));
    }, [slides.length]);

    const goToSlide = React.useCallback((index: number) => {
        setCurrentIndex(index);
    }, []);

    const stopTimer = React.useCallback(() => {
        if (timerRef.current) {
            clearTimeout(timerRef.current);
            timerRef.current = null;
        }
    }, []);

    const startTimer = React.useCallback(() => {
        stopTimer();
        if (slides.length > 1) {
            timerRef.current = setTimeout(() => {
                nextSlide();
            }, props.autoRotateInterval * 1000);
        }
    }, [slides.length, props.autoRotateInterval, nextSlide, stopTimer]);

    const fetchData = React.useCallback(async () => {
        if (!props.siteUrl || !props.listId || !props.titleColumn || !props.imageColumn) {
            setSlides([]);
            setIsLoading(false);
            return;
        }

        try {
            setIsLoading(true);
            setError(null);

            const apiUrl = `${props.siteUrl}/_api/web/lists(guid'${props.listId}')/renderListDataAsStream`;

            // Construct the select columns
            // (selectColumns is logged via data.Row structure in renderListDataAsStream)

            const body = {
                parameters: {
                    RenderOptions: 2, // ListData
                    ViewXml: `<View><Query><OrderBy><FieldRef Name="Title" Ascending="TRUE"/></OrderBy></Query></View>`
                }
            };

            const response = await props.context.spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, {
                body: JSON.stringify(body)
            });

            if (!response.ok) {
                const error = await response.text();
                throw new Error(`Error fetching data: ${response.status} - ${error}`);
            }

            const data = await response.json();

            const mappedSlides: ISlideProps[] = (data.Row || []).map((item: ISharePointRow, index: number) => {
                const imageUrl = getImageUrl(item, props.imageColumn, props.siteUrl, props.siteId, props.webId);

                return {
                    title: item[props.titleColumn] || '',
                    description: item[props.descriptionColumn] || '',
                    imageUrl: imageUrl,
                    buttonText: item[props.buttonTextColumn] || '',
                    pageLink: item[props.pageLinkColumn] || '',
                    showCta: props.showCta,
                    isActive: index === 0
                };
            });

            setSlides(mappedSlides);
            setCurrentIndex(0);
        } catch (err) {
            setError((err as Error).message || 'Error loading slides');
        } finally {
            setIsLoading(false);
        }
    }, [props]);

    useEffect(() => {
        fetchData().catch(e => console.error(e));
    }, [fetchData]);

    useEffect(() => {
        startTimer();
        return () => stopTimer();
    }, [currentIndex, slides.length, props.autoRotateInterval, startTimer, stopTimer]);

    if (isLoading) {
        return <div className={styles.bannerSlider}><EmptyState message="Loading slides..." /></div>;
    }

    if (error) {
        return <div className={styles.bannerSlider}><EmptyState message={`Error: ${error}`} /></div>;
    }

    if (!slides || slides.length === 0) {
        return (
            <section
                className={styles.bannerSlider}
                aria-label="Banner Slider"
            >
                <EmptyState
                    icon="PhotoCollection"
                    message="No slides to display. Please configure the data source and column mappings in the web part properties."
                />
            </section>
        );
    }

    return (
        <section
            className={styles.bannerSliderContainer}
            aria-roledescription="carousel"
            aria-label="Promotional Banners"
        >
            {props.showTitle && props.title && (
                <div className={styles.webpartHeader}>
                    <div className={styles.titleContainer}>
                        <h2>{props.title}</h2>
                        {props.showBackgroundBar && <div className={styles.backgroundBar} />}
                    </div>
                </div>
            )}
            <div
                className={styles.bannerSlider}
                onMouseEnter={stopTimer}
                onMouseLeave={startTimer}
                onFocus={stopTimer}
                onBlur={startTimer}
                tabIndex={0}
            >
                <div
                    className={styles.sliderContent}
                    style={{
                        transform: `translateX(-${currentIndex * (100 / slides.length)}%)`,
                        width: `${slides.length * 100}%`
                    }}
                >
                    {slides.map((slide, index) => {
                        // Fix for server relative URLs and origin
                        let finalImageUrl = slide.imageUrl;
                        if (finalImageUrl?.startsWith('/')) {
                            const origin = globalThis.location.origin;
                            finalImageUrl = `${origin}${finalImageUrl}`;
                        }
                        if (finalImageUrl?.startsWith('/')) {
                            const origin = globalThis.location.origin;
                            finalImageUrl = `${origin}${finalImageUrl}`;
                        }

                        const slideKey = finalImageUrl ? `${slide.title}-${finalImageUrl}` : index;
                        return (
                            <div key={slideKey as React.Key} style={{ width: `${100 / slides.length}%` }}>
                                <Slide
                                    {...slide}
                                    imageUrl={finalImageUrl}
                                    isActive={index === currentIndex}
                                    showCta={props.showCta}
                                />
                            </div>
                        );
                    })}
                </div>

                {slides.length > 1 && (
                    <>
                        <button
                            className={`${styles.navButton} ${styles.prevButton}`}
                            onClick={prevSlide}
                            aria-label="Previous slide"
                        >
                            <svg viewBox="0 0 24 24"><path d="M15.41 7.41L14 6l-6 6 6 6 1.41-1.41L10.83 12z" /></svg>
                        </button>
                        <button
                            className={`${styles.navButton} ${styles.nextButton}`}
                            onClick={nextSlide}
                            aria-label="Next slide"
                        >
                            <svg viewBox="0 0 24 24"><path d="M10 6L8.59 7.41 13.17 12l-4.58 4.59L10 18l6-6z" /></svg>
                        </button>

                        <div className={styles.dotsContainer}>
                            {slides.map((slide, index) => {
                                const dotKey = slide.imageUrl ? `dot-${slide.title}-${slide.imageUrl}` : `dot-${index}`;
                                return (
                                    <button
                                        key={dotKey as React.Key}
                                        className={`${styles.dot} ${index === currentIndex ? styles.active : ''}`}
                                        onClick={() => goToSlide(index)}
                                        aria-label={`Go to slide ${index + 1}`}
                                    />
                                );
                            })}
                        </div>
                    </>
                )}
            </div>
        </section>
    );
};
