import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import styles from './NewJoiners.module.scss';
import { INewJoinersProps } from './INewJoinersProps';
import { NewJoinerCard, INewJoiner } from './NewJoinerCard';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { EmptyState } from '../../../common/components/EmptyState/EmptyState';

export const NewJoiners: React.FC<INewJoinersProps> = (props) => {
    const [joiners, setJoiners] = useState<INewJoiner[]>([]);
    const [loading, setLoading] = useState<boolean>(true);
    const [currentSlide, setCurrentSlide] = useState<number>(0);
    const [isHovering, setIsHovering] = useState<boolean>(false);
    const timerRef = useRef<any>(null);

    const handleSlide = (direction: number): void => {
        if (joiners.length <= 1) return;
        setCurrentSlide(prev => {
            let next = (prev + direction) % joiners.length;
            if (next < 0) next = joiners.length - 1;
            return next;
        });
    };

    const goToSlide = (index: number): void => {
        setCurrentSlide(index);
    };

    const loadGraphData = async (): Promise<void> => {
        if (!props.manualJoiners || props.manualJoiners.length === 0) {
            setJoiners([]);
            return;
        }

        const mappedJoiners: INewJoiner[] = props.manualJoiners.map(mj => ({
            name: mj.user?.fullName || mj.user?.text || 'Unknown',
            jobTitle: mj.user?.jobTitle || mj.user?.secondaryText || '',
            department: mj.user?.department || '',
            photoUrl: mj.user?.imageUrl || `/_layouts/15/userphoto.aspx?size=L&accountname=${mj.user?.loginName || mj.user?.secondaryText}`,
            introText: mj.introText || props.commonIntro,
            email: mj.user?.loginName || mj.user?.secondaryText
        }));

        setJoiners(mappedJoiners);
    };

    const loadSPData = async (): Promise<void> => {
        if (!props.siteUrl || !props.listId || !props.newJoinerColumn) {
            setJoiners([]);
            return;
        }

        const selectFields = [
            'Id',
            props.nameColumn,
            props.photoColumn,
            props.jobTitleColumn,
            props.departmentColumn,
            props.newJoinerColumn,
            props.newJoinerTextColumn
        ].filter(f => !!f);

        if (props.emailColumn) {
            selectFields.push(`${props.emailColumn}/Title`, `${props.emailColumn}/EMail`);
        }

        const expand = props.emailColumn ? `&$expand=${props.emailColumn}` : '';
        const filter = `${props.newJoinerColumn} eq 1`;
        const top = props.maxItems || 5;

        const endpoint = `${props.siteUrl}/_api/web/lists(guid'${props.listId}')/items?$select=${selectFields.join(',')}${expand}&$filter=${filter}&$orderby=Created desc&$top=${top}`;

        const response: SPHttpClientResponse = await props.context.spHttpClient.get(
            endpoint,
            SPHttpClient.configurations.v1
        );

        if (!response.ok) {
            throw new Error(`Error fetching New Joiners: ${response.statusText}`);
        }

        const data = await response.json();
        const items = data.value || [];

        const mappedJoiners: INewJoiner[] = items.map((item: any) => {
            const photoData = item[props.photoColumn];
            let photoUrl = '';
            if (props.photoColumn && photoData) {
                if (typeof photoData === 'string') {
                    try {
                        const parsed = JSON.parse(photoData);
                        photoUrl = parsed.serverRelativeUrl || parsed.src || photoData;
                    } catch {
                        photoUrl = photoData;
                    }
                } else {
                    photoUrl = photoData.serverRelativeUrl || photoData.src || '';
                }

                if (photoUrl && !photoUrl.startsWith('http')) {
                    photoUrl = `${new URL(props.siteUrl).origin}${photoUrl}`;
                }
            }

            const emailInfo = props.emailColumn ? item[props.emailColumn]?.EMail : '';

            return {
                name: item[props.nameColumn] || 'Unknown',
                jobTitle: item[props.jobTitleColumn] || '',
                department: item[props.departmentColumn] || '',
                photoUrl: photoUrl,
                introText: props.newJoinerTextColumn ? item[props.newJoinerTextColumn] : '',
                email: emailInfo
            };
        });

        setJoiners(mappedJoiners);
    };

    const loadData = async (): Promise<void> => {
        setLoading(true);
        try {
            if (props.source === 'graph') {
                await loadGraphData();
            } else {
                await loadSPData();
            }
        } catch (error) {
            console.error('Error loading New Joiners data:', error);
            setJoiners([]);
        } finally {
            setLoading(false);
        }
    };

    useEffect(() => {
        loadData().catch(console.error);
    }, [
        props.source,
        props.siteUrl,
        props.listId,
        props.newJoinerColumn,
        props.nameColumn,
        props.jobTitleColumn,
        props.departmentColumn,
        props.emailColumn,
        props.maxItems,
        props.manualJoiners,
        props.commonIntro
    ]);

    useEffect(() => {
        if (!isHovering && joiners.length > 1 && props.autoRotateInterval && props.layout === 'strip') {
            timerRef.current = setTimeout(() => {
                handleSlide(1);
            }, props.autoRotateInterval * 1000);
        }
        return () => {
            if (timerRef.current) {
                clearTimeout(timerRef.current);
            }
        };
    }, [currentSlide, joiners.length, isHovering, props.autoRotateInterval, props.layout]);

    const isConfigured = props.source === 'graph'
        ? (props.manualJoiners && props.manualJoiners.length > 0)
        : (props.siteUrl && props.listId && props.nameColumn && props.jobTitleColumn && props.departmentColumn && props.newJoinerColumn);

    if (!isConfigured) {
        return (
            <div className={styles.newJoiners}>
                <EmptyState
                    icon="People"
                    title="New Joiners - Configuration Required"
                    message="Please complete the web part configuration to display new joiners."
                    description="Specify the data source and map the required columns in the property pane."
                />
            </div>
        );
    }

    const containerClass = `${styles.newJoiners} ${props.layoutMode === 'compact' ? styles.compact : ''}`;

    function renderHeader(): JSX.Element | null {
        if (!props.webPartTitle) return null;

        const isSolid = props.titleBarStyle === 'solid';
        const headerClass = props.showBackgroundBar
            ? (isSolid ? styles.solidBackground : styles.underlineBackground)
            : '';

        return (
            <div className={`${styles.webpartHeader} ${headerClass}`}>
                <div className={styles.titleContainer}>
                    <h2>{props.webPartTitle}</h2>
                </div>
            </div>
        );
    }

    const renderSlider = (): JSX.Element => (
        <div
            className={styles.strip}
            onMouseEnter={() => setIsHovering(true)}
            onMouseLeave={() => setIsHovering(false)}
            role="region"
            aria-label="New Joiners Carousel"
        >
            <div
                className={styles.sliderTrack}
                style={{ transform: `translateX(-${currentSlide * 100}%)` }}
            >
                {joiners.map((joiner, index) => (
                    <div key={`${joiner.email || joiner.name}-${index}`} className={styles.sliderItem}>
                        <NewJoinerCard
                            joiner={joiner}
                            layout="strip"
                        />
                    </div>
                ))}
            </div>

            {joiners.length > 1 && (
                <>
                    <button
                        type="button"
                        className={`${styles.sliderArrow} ${styles.left}`}
                        onClick={() => handleSlide(-1)}
                        aria-label="Previous Joiner"
                    >
                        <svg viewBox="0 0 24 24"><path d="M15.41 7.41L14 6l-6 6 6 6 1.41-1.41L10.83 12z" /></svg>
                    </button>
                    <button
                        type="button"
                        className={`${styles.sliderArrow} ${styles.right}`}
                        onClick={() => handleSlide(1)}
                        aria-label="Next Joiner"
                    >
                        <svg viewBox="0 0 24 24"><path d="M10 6L8.59 7.41 13.17 12l-4.58 4.59L10 18l6-6z" /></svg>
                    </button>
                    <div className={styles.sliderDots}>
                        {joiners.map((joiner, index) => (
                            <button
                                key={`${joiner.email || joiner.name}-${index}`}
                                type="button"
                                className={`${styles.sliderDot} ${index === currentSlide ? styles.active : ''}`}
                                onClick={() => goToSlide(index)}
                                aria-label={`Go to joiner ${index + 1}`}
                            />
                        ))}
                    </div>
                </>
            )}
        </div>
    );

    const renderContent = (): JSX.Element => {
        if (loading) {
            return (
                <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '200px' }}>
                    <Spinner size={SpinnerSize.large} label="Loading new joiners..." />
                </div>
            );
        }

        if (joiners.length === 0) {
            return (
                <EmptyState
                    icon="ContactInfo"
                    title="No New Joiners"
                    message="There are no recent new joiners to display."
                    description="Add items to your SharePoint list or check your filter settings if applicable."
                />
            );
        }

        if (props.layout === 'strip') {
            return renderSlider();
        }

        return (
            <div className={`${styles.grid} ${props.layout === 'list' ? styles.list : styles.grid}`}>
                {joiners.map((joiner, index) => (
                    <NewJoinerCard
                        key={`${joiner.email || joiner.name}-${index}`}
                        joiner={joiner}
                        layout={props.layout}
                    />
                ))}
            </div>
        );
    };

    return (
        <section className={containerClass}>
            {renderHeader()}
            <div className={styles.content}>
                {renderContent()}
            </div>
        </section>
    );
};
