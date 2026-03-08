/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { ICalendarProps, ICalendarEvent } from './ICalendarProps';
import styles from './Calendar.module.scss';
import CalendarHeader from './CalendarHeader';
import { useCalendarNavigation } from './useCalendarNavigation';
import { CalendarService } from './CalendarService';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { EmptyState } from '../../../common/components/EmptyState/EmptyState';

import CalendarMonthView from './CalendarMonthView';
import CalendarDayView from './CalendarDayView';
import CalendarWeekView from './CalendarWeekView';
import CalendarYearView from './CalendarYearView';

const TITLE_STYLE_SOLID = 'solid';
const TITLE_STYLE_UNDERLINE = 'underline';

export const Calendar: React.FC<ICalendarProps> = (props) => {
    const {
        siteUrl, listId,
        titleColumn, dateColumn, endDateColumn, locationColumn,
        defaultView, context,
        showTitle, title, showBackgroundBar, titleBarStyle,
        yearViewType
    } = props;

    const {
        currentDate,
        view,
        navigate,
        snapToToday,
        changeView,
        getViewRange
    } = useCalendarNavigation(defaultView);

    const [events, setEvents] = React.useState<ICalendarEvent[]>([]);
    const [loading, setLoading] = React.useState<boolean>(false);
    const calendarService = React.useMemo(() => new CalendarService(context), [context]);

    const fetchEvents = React.useCallback(async () => {
        if (!siteUrl || !listId || !titleColumn || !dateColumn) return;

        setLoading(true);
        try {
            const range = getViewRange();
            if (!range) return;

            const results = await calendarService.getSPListEvents(
                siteUrl,
                listId,
                range.start,
                range.end,
                {
                    title: titleColumn,
                    date: dateColumn,
                    endDate: endDateColumn,
                    location: locationColumn
                }
            );
            setEvents(results);
        } catch (error) {
            console.error('Error fetching calendar events:', error);
        } finally {
            setLoading(false);
        }
    }, [siteUrl, listId, titleColumn, dateColumn, endDateColumn, locationColumn, getViewRange, calendarService]);

    React.useEffect(() => {
        fetchEvents().catch(console.error);
    }, [fetchEvents]);

    const getHeaderClass = (): string => {
        if (!showBackgroundBar) return '';
        return titleBarStyle === TITLE_STYLE_SOLID ? styles.solidBackground : styles.underlineBackground;
    };

    const renderView = (): JSX.Element => {
        const viewProps = { currentDate, events, loading, yearViewType };
        switch (view) {
            case 'day': return <CalendarDayView {...viewProps} />;
            case 'week': return <CalendarWeekView {...viewProps} />;
            case 'month': return <CalendarMonthView {...viewProps} />;
            case 'year': return <CalendarYearView {...viewProps} />;
            default: return <CalendarMonthView {...viewProps} />;
        }
    };

    if (!siteUrl || !listId || !titleColumn || !dateColumn) {
        return (
            <div className={styles.calendar}>
                <EmptyState
                    icon="Calendar"
                    title="Calendar - Configuration Required"
                    message="Please complete the web part configuration to display events."
                    description="You need to specify the Site URL, List ID, and map the required columns (Title and Start Date) in the property pane."
                />
            </div>
        );
    }

    if (loading) {
        return (
            <div className={styles.calendar}>
                {showTitle && title && (
                    <div className={`${styles.webpartHeader} ${getHeaderClass()}`}>
                        <div className={styles.titleContainer}>
                            <h2>{title}</h2>
                        </div>
                    </div>
                )}
                <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '300px' }}>
                    <Spinner size={SpinnerSize.large} label="Loading calendar events..." />
                </div>
            </div>
        );
    }

    return (
        <div className={styles.calendar}>
            {showTitle && title && (
                <div className={`${styles.webpartHeader} ${getHeaderClass()}`}>
                    <div className={styles.titleContainer}>
                        <h2>{title}</h2>
                    </div>
                </div>
            )}
            <CalendarHeader
                currentDate={currentDate}
                view={view}
                onNavigate={navigate}
                onToday={snapToToday}
                onViewChange={changeView}
            />
            <div className={styles.viewBody}>
                {renderView()}
            </div>
        </div>
    );
};

export default Calendar;
