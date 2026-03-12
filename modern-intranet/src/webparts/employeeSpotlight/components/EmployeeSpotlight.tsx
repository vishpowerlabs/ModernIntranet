/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { IEmployeeSpotlightProps } from './IEmployeeSpotlightProps';
import { SpotlightCard, ISpotlightItem } from './SpotlightCard';
import styles from './EmployeeSpotlight.module.scss';
import { EmptyState } from '../../../common/components/EmptyState/EmptyState';
import { WebPartHeader } from '../../../common/components/WebPartHeader/WebPartHeader';

const EmployeeSpotlight: React.FC<IEmployeeSpotlightProps> = (props) => {
    const [loading, setLoading] = useState<boolean>(true);
    const [error, setError] = useState<string | null>(null);
    const [spotlights, setSpotlights] = useState<ISpotlightItem[]>([]);
    const [currentSlide, setCurrentSlide] = useState<number>(0);
    const [isHovering, setIsHovering] = useState<boolean>(false);
    const timerRef = useRef<number | null>(null);

    const handleSlide = (direction: number): void => {
        if (spotlights.length <= 1) return;
        setCurrentSlide(prev => {
            let next = (prev + direction) % spotlights.length;
            if (next < 0) next = spotlights.length - 1;
            return next;
        });
    };

    const loadGraphData = (): void => {
        if (!props.selectedUsers || props.selectedUsers.length === 0) {
            setSpotlights([]);
            return;
        }

        const mappedItems: ISpotlightItem[] = (props.selectedUsers || []).map((user: any) => ({
            Id: user.id || user.secondaryText,
            Name: user.text || '',
            JobTitle: user.tertiaryText || '',
            Department: user.optionalText || '',
            SpotlightText: props.commonDescription || '',
            Email: user.secondaryText,
            PhotoUrl: user.imageUrl
        }));
        setSpotlights(mappedItems);
    };

    const loadSPData = async (): Promise<void> => {
        if (!props.siteUrl || !props.listId || !props.nameColumn || !props.jobTitleColumn || !props.departmentColumn || !props.spotlightColumn || !props.spotlightTextColumn) {
            return;
        }

        let selectFields = `Id,${props.nameColumn},${props.jobTitleColumn},${props.departmentColumn},${props.spotlightColumn},${props.spotlightTextColumn}`;
        let expandFields = '';

        if (props.photoColumn) {
            selectFields += `,${props.photoColumn}`;
        }

        if (props.emailColumn) {
            selectFields += `,${props.emailColumn}/Title,${props.emailColumn}/EMail`;
            expandFields = `&$expand=${props.emailColumn}`;
        }

        const apiUrl = `${props.siteUrl}/_api/web/lists(guid'${props.listId}')/items?$select=${selectFields}${expandFields}&$filter=${props.spotlightColumn} eq 1&$orderby=Modified desc&$top=${props.maxItems || 3}`;

        const response: SPHttpClientResponse = await props.context.spHttpClient.get(
            apiUrl,
            SPHttpClient.configurations.v1
        );

        if (!response.ok) {
            throw new Error(`Data fetch failed: ${response.statusText}`);
        }

        const data = await response.json();

        const mappedItems: ISpotlightItem[] = data.value.map((item: any) => {
            let photoUrl: string | undefined = undefined;
            if (props.photoColumn && item[props.photoColumn]) {
                try {
                    const photoObj = JSON.parse(item[props.photoColumn]);
                    photoUrl = photoObj.serverRelativeUrl;
                } catch {
                    photoUrl = item[props.photoColumn];
                }
            }

            return {
                Id: item.Id,
                Name: item[props.nameColumn] || '',
                JobTitle: item[props.jobTitleColumn] || '',
                Department: item[props.departmentColumn] || '',
                SpotlightText: item[props.spotlightTextColumn] || '',
                Email: item[props.emailColumn]?.EMail,
                PhotoUrl: photoUrl
            };
        });

        setSpotlights(mappedItems);
    };

    const loadData = async (): Promise<void> => {
        try {
            setLoading(true);
            setError(null);

            if (props.source === 'graph') {
                loadGraphData();
            } else {
                await loadSPData();
            }
            setCurrentSlide(0);
        } catch (err) {
            setError(err instanceof Error ? err.message : String(err));
        } finally {
            setLoading(false);
        }
    };

    const goToSlide = (index: number): void => {
        setCurrentSlide(index);
    };

    useEffect(() => {
        loadData().catch((err) => {
            console.error('Error in EmployeeSpotlight loadData:', err);
        });
    }, [props.source, props.siteUrl, props.listId, props.maxItems, props.spotlightColumn, props.selectedUsers, props.commonDescription, props.layoutMode]);

    useEffect(() => {
        if (spotlights.length > 1 && !isHovering) {
            timerRef.current = globalThis.setTimeout(() => {
                handleSlide(1);
            }, (props.autoRotateInterval || 5) * 1000) as unknown as number;
        }
        return () => {
            if (timerRef.current) {
                clearTimeout(timerRef.current);
            }
        };
    }, [currentSlide, spotlights.length, isHovering, props.autoRotateInterval]);

    const isConfigured = props.source === 'graph'
        ? (props.selectedUsers && props.selectedUsers.length > 0)
        : (props.siteUrl && props.listId && props.nameColumn && props.jobTitleColumn && props.departmentColumn && props.spotlightColumn && props.spotlightTextColumn);

    const renderHeader = (): JSX.Element => (
        <WebPartHeader
            title={props.title || ''}
            showTitle={props.showTitle}
            showBackgroundBar={!!props.showBackgroundBar}
            titleBarStyle={props.titleBarStyle || 'underline'}
        />
    );

    if (!isConfigured) {
        return (
            <div className={styles.spotlightContainer}>
                <EmptyState
                    icon="Contact"
                    title="Employee Spotlight - Configuration Required"
                    message="Please complete the web part configuration to display content."
                    description="You need to specify the Site URL, List ID, and map the required columns (Name, Job Title, Department, Spotlight Toggle, and Spotlight Text) in the property pane."
                />
            </div>
        );
    }

    if (loading) {
        return (
            <div className={styles.spotlightContainer}>
                {renderHeader()}
                <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '300px' }}>
                    <Spinner size={SpinnerSize.large} label="Loading spotlights..." />
                </div>
            </div>
        );
    }

    if (error) {
        return (
            <div className={styles.spotlightContainer}>
                <EmptyState icon="Error" message={error} />
            </div>
        );
    }

    if (spotlights.length === 0) {
        return (
            <div className={styles.spotlightContainer}>
                <EmptyState
                    icon="People"
                    title="No Spotlighted Employees"
                    message="There are no employees marked for spotlight display."
                    description="Set the Spotlight column to 'Yes' for employee records in your target list to show them here."
                />
            </div>
        );
    }

    return (
        <div className={`${styles.spotlightContainer} ${props.layoutMode === 'compact' ? styles.compact : ''}`}>
            {renderHeader()}

            <section
                className={styles.wpSpotlight}
                onMouseEnter={() => setIsHovering(true)}
                onMouseLeave={() => setIsHovering(false)}
                aria-label="Employee Spotlight Carousel"
            >
                <div
                    className={styles.wpSpotlightTrack}
                    style={{ transform: `translateX(-${currentSlide * 100}%)` }}
                >
                    {spotlights.map(item => (
                        <SpotlightCard key={item.Id} item={item} />
                    ))}
                </div>

                {spotlights.length > 1 && (
                    <>
                        <button
                            type="button"
                            className={`${styles.wpSpotlightArrow} ${styles.left}`}
                            onClick={(e) => { e.preventDefault(); handleSlide(-1); }}
                            aria-label="Previous Highlight"
                        >
                            <svg viewBox="0 0 24 24">
                                <path d="M15.41 7.41L14 6l-6 6 6 6 1.41-1.41L10.83 12z" />
                            </svg>
                        </button>
                        <button
                            type="button"
                            className={`${styles.wpSpotlightArrow} ${styles.right}`}
                            onClick={(e) => { e.preventDefault(); handleSlide(1); }}
                            aria-label="Next Highlight"
                        >
                            <svg viewBox="0 0 24 24">
                                <path d="M10 6L8.59 7.41 13.17 12l-4.58 4.59L10 18l6-6z" />
                            </svg>
                        </button>
                        <div className={styles.wpSpotlightDots}>
                            {spotlights.map((_, index) => (
                                <button
                                    key={spotlights[index].Id}
                                    type="button"
                                    className={`${styles.wpSpotlightDot} ${index === currentSlide ? styles.active : ''}`}
                                    onClick={() => goToSlide(index)}
                                    aria-label={`Go to item ${index + 1}`}
                                />
                            ))}
                        </div>
                    </>
                )}
            </section>
        </div >
    );
};

export default EmployeeSpotlight;
