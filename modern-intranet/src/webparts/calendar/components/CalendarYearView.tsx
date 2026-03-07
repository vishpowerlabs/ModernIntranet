/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import {
    format,
    startOfYear,
    eachMonthOfInterval,
    endOfYear,
    startOfMonth,
    endOfMonth,
    eachDayOfInterval,
    startOfWeek,
    endOfWeek,
    isSameDay,
    isToday
} from 'date-fns';
import { ICalendarEvent } from './ICalendarProps';
import styles from './Calendar.module.scss';

interface ICalendarYearViewProps {
    currentDate: Date;
    events: ICalendarEvent[];
}

interface ITimelineDayCellProps {
    month: Date;
    day: number;
    daysInMonth: number;
    events: ICalendarEvent[];
    styles: any;
}

const TimelineDayCell: React.FC<ITimelineDayCellProps> = ({ month, day, daysInMonth, events, styles }) => {
    const isDateValid = day <= daysInMonth;
    const date = new Date(month.getFullYear(), month.getMonth(), day);
    const hasEvents = isDateValid && events.some(ev => isSameDay(ev.start, date));
    const isCurrentDay = isDateValid && isToday(date);

    return (
        <div
            className={`
                ${styles.dayCell} 
                ${isDateValid ? '' : styles.invalidDay}
                ${hasEvents ? styles.hasEvents : ''}
                ${isCurrentDay ? styles.today : ''}
            `}
            title={hasEvents ? `${format(date, 'MMM dd')}: ${events.filter(ev => isSameDay(ev.start, date)).map(ev => ev.title).join(', ')}` : ''}
        />
    );
};

const CalendarYearView: React.FC<ICalendarYearViewProps & { yearViewType?: 'grid' | 'timeline' }> = ({ currentDate, events, yearViewType = 'grid' }): JSX.Element => {
    const yearStart = startOfYear(currentDate);
    const yearEnd = endOfYear(yearStart);
    const months = eachMonthOfInterval({ start: yearStart, end: yearEnd });

    const renderMiniMonth = (monthDate: Date): JSX.Element => {
        const monthStart = startOfMonth(monthDate);
        const monthEnd = endOfMonth(monthStart);
        const startDate = startOfWeek(monthStart);
        const endDate = endOfWeek(monthEnd);
        const days = eachDayOfInterval({ start: startDate, end: endDate });

        return (
            <div key={monthDate.toISOString()} className={styles.miniMonth}>
                <div className={styles.miniMonthName}>{format(monthDate, 'MMMM')}</div>
                <div className={styles.miniMonthGrid}>
                    {['S', 'M', 'T', 'W', 'T', 'F', 'S'].map((d, i) => (
                        <div key={`${monthDate.toISOString()}-header-${i}`} className={styles.miniDayHeader}>{d}</div>
                    ))}
                    {days.map((day) => {
                        const hasEvents = events.some(ev => isSameDay(ev.start, day));
                        const isCurrMonth = day.getMonth() === monthDate.getMonth();

                        return (
                            <div
                                key={day.toISOString()}
                                className={`
                  ${styles.miniDay} 
                  ${isCurrMonth ? '' : styles.miniOtherMonth} 
                  ${hasEvents ? styles.hasEvents : ''}
                  ${isToday(day) ? styles.today : ''}
                `}
                            >
                                {format(day, 'd')}
                            </div>
                        );
                    })}
                </div>
            </div>
        );
    };

    const renderTimelineView = (): JSX.Element => {
        return (
            <div className={styles.yearTimeline}>
                <div className={styles.timelineHeader}>
                    <div className={styles.monthCol}>Month</div>
                    <div className={styles.daysContainer}>
                        {Array.from({ length: 31 }, (_, i) => (
                            <div key={`day-head-${i + 1}`} className={styles.dayHead}>{i + 1}</div>
                        ))}
                    </div>
                </div>
                {months.map(month => (
                    <div key={month.toISOString()} className={styles.monthRow}>
                        <div className={styles.monthName}>{format(month, 'MMM')}</div>
                        <div className={styles.daysContainer}>
                            {Array.from({ length: 31 }, (_, i) => (
                                <TimelineDayCell
                                    key={`day-cell-${month.getMonth()}-${i + 1}`}
                                    month={month}
                                    day={i + 1}
                                    daysInMonth={endOfMonth(month).getDate()}
                                    events={events}
                                    styles={styles}
                                />
                            ))}
                        </div>
                    </div>
                ))}
            </div>
        );
    };

    return (
        <div className={yearViewType === 'timeline' ? styles.yearTimelineContainer : styles.yearView}>
            {yearViewType === 'timeline' ? renderTimelineView() : months.map(m => renderMiniMonth(m))}
        </div>
    );
};

export default CalendarYearView;
