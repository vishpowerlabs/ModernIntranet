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
import { EmptyState } from '../../../common/components/EmptyState/EmptyState';

import CalendarMonthView from './CalendarMonthView';
import CalendarDayView from './CalendarDayView';
import CalendarWeekView from './CalendarWeekView';
import CalendarYearView from './CalendarYearView';

const Calendar: React.FC<ICalendarProps> = (props) => {
    const {
        siteUrl, listId,
        titleColumn, dateColumn, endDateColumn, locationColumn,
        defaultView, context,
        showTitle, title, showBackgroundBar,
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
        void fetchEvents();
    }, [fetchEvents]);

    if (!siteUrl || !listId || !titleColumn || !dateColumn) {
        return (
            <EmptyState
                message="Please configure the calendar data source and column mappings in the property pane."
                icon="Calendar"
            />
        );
    }

    const renderView = (): JSX.Element => {
        const viewProps = { currentDate, events, loading, yearViewType };
        if (loading) {
            return (
                <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '100%' }}>
                    Loading...
                </div>
            );
        }
        switch (view) {
            case 'day': return <CalendarDayView {...viewProps} />;
            case 'week': return <CalendarWeekView {...viewProps} />;
            case 'month': return <CalendarMonthView {...viewProps} />;
            case 'year': return <CalendarYearView {...viewProps} />;
        }
    };

    return (
        <div className={styles.calendar}>
            {showTitle && title && (
                <div className={styles.webpartHeader}>
                    <div className={styles.titleContainer}>
                        <h2>{title}</h2>
                        {showBackgroundBar && <div className={styles.backgroundBar} />}
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
