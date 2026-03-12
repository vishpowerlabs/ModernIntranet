/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import styles from './BannerSlider.module.scss';
import { IBannerSliderProps } from './IBannerSliderProps';
import { Slide, ISlideProps } from './Slide';
import { EmptyState } from '../../../common/components/EmptyState/EmptyState';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartHeader } from '../../../common/components/WebPartHeader/WebPartHeader';

interface ISharePointImageMetadata {
    fileName?: string;
    serverRelativeUrl?: string;
    serverUrl?: string;
    Url?: string;
}

interface ISharePointRow {
    ID?: string;
    FileDirRef?: string;
    [key: string]: string | number | boolean | undefined | null;
}

const getPreviewUrl = (fileName: string, siteUrl: string, siteId: string, webId: string): string => {
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
    const [loading, setLoading] = useState(true);
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
            setLoading(false);
            return;
        }

        try {
            setLoading(true);
            setError(null);

            const apiUrl = `${props.siteUrl}/_api/web/lists(guid'${props.listId}')/renderListDataAsStream`;

            const body = {
                parameters: {
                    RenderOptions: 2,
                    ViewXml: `<View><Query><OrderBy><FieldRef Name="Title" Ascending="TRUE"/></OrderBy></Query></View>`
                }
            };

            const response = await props.context.spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, {
                body: JSON.stringify(body)
            });

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Error fetching data: ${response.status} - ${errorText}`);
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
            setLoading(false);
        }
    }, [props]);

    useEffect(() => {
        fetchData().catch(e => console.error(e));
    }, [fetchData]);

    useEffect(() => {
        startTimer();
        return () => stopTimer();
    }, [currentIndex, slides.length, props.autoRotateInterval, startTimer, stopTimer]);

    const isConfigured = props.siteUrl && props.listId && props.titleColumn && props.imageColumn;

    const renderHeader = (): JSX.Element => (
        <WebPartHeader
            title={props.title || ''}
            showTitle={!!props.showTitle}
            showBackgroundBar={!!props.showBackgroundBar}
            titleBarStyle={props.titleBarStyle || 'underline'}
        />
    );

    if (!isConfigured) {
        return (
            <section className={styles.bannerSliderContainer}>
                <EmptyState
                    icon="PhotoCollection"
                    title="Banner Slider Web Part - Configuration Required"
                    message="Please complete the web part configuration to display content."
                    description="You need to specify the Site URL, List ID, and map the required columns (Title and Image) in the property pane."
                />
            </section>
        );
    }

    if (loading) {
        return (
            <div className={styles.bannerSliderContainer}>
                {renderHeader()}
                <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '300px' }}>
                    <Spinner size={SpinnerSize.large} label="Loading promotional banners..." />
                </div>
            </div>
        );
    }

    if (error) {
        return (
            <div className={styles.bannerSliderContainer}>
                {renderHeader()}
                <EmptyState
                    icon="Error"
                    title="Error Loading Banners"
                    message={error}
                    description="Please verify your Site URL and List ID in the property pane."
                />
            </div>
        );
    }

    if (!slides || slides.length === 0) {
        return (
            <section className={styles.bannerSliderContainer}>
                {renderHeader()}
                <EmptyState
                    icon="PhotoCollection"
                    title="No Slides Found"
                    message="There are no active slides to display from the selected list."
                    description="Add items to your SharePoint list or check your filter settings if applicable."
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
            {renderHeader()}
            <div
                className={styles.bannerSlider}
                onMouseEnter={stopTimer}
                onMouseLeave={startTimer}
                role="region"
                aria-label="Promotional Banners Carousel"
            >
                <div
                    className={styles.sliderContent}
                    style={{
                        transform: `translateX(-${currentIndex * (100 / slides.length)}%)`,
                        width: `${slides.length * 100}%`
                    }}
                >
                    {slides.map((slide, index) => {
                        let finalImageUrl = slide.imageUrl;
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
                            {slides.map((_slide, index) => {
                                return (
                                    <button
                                        key={`dot-${index}`}
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
        </section >
    );
};
