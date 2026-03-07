/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import {
    format,
    startOfMonth,
    endOfMonth,
    startOfWeek,
    endOfWeek,
    eachDayOfInterval,
    isSameMonth,
    isSameDay,
    isToday
} from 'date-fns';
import { ICalendarEvent } from './ICalendarProps';
import styles from './Calendar.module.scss';

interface ICalendarMonthViewProps {
    currentDate: Date;
    events: ICalendarEvent[];
}

const CalendarMonthView: React.FC<ICalendarMonthViewProps> = ({ currentDate, events }): JSX.Element => {
    const monthStart = startOfMonth(currentDate);
    const monthEnd = endOfMonth(monthStart);
    const startDate = startOfWeek(monthStart);
    const endDate = endOfWeek(monthEnd);

    const days = eachDayOfInterval({ start: startDate, end: endDate });
    const weekDays = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

    const getEventsForDay = (day: Date): ICalendarEvent[] => {
        return events.filter(ev => isSameDay(ev.start, day));
    };

    return (
        <div className={styles.monthView}>
            <div className={styles.monthHeader}>
                {weekDays.map(d => <div key={d} className={styles.dayHeader}>{d}</div>)}
            </div>
            <div className={styles.monthGrid}>
                {days.map((day) => {
                    const dayEvents = getEventsForDay(day);
                    const isCurrentMonth = isSameMonth(day, monthStart);

                    return (
                        <div
                            key={day.toISOString()}
                            className={`${styles.monthCell} ${isCurrentMonth ? '' : styles.otherMonth}`}
                        >
                            <div className={styles.cellHeader}>
                                <span className={`${styles.dayNumber} ${isToday(day) ? styles.today : ''}`}>
                                    {format(day, 'd')}
                                </span>
                            </div>
                            <div className={styles.cellEvents}>
                                {dayEvents.slice(0, 3).map(ev => (
                                    <div key={ev.id} className={styles.eventPill} title={ev.title}>
                                        {ev.title}
                                    </div>
                                ))}
                                {dayEvents.length > 3 && (
                                    <div className={styles.moreLink}>
                                        +{dayEvents.length - 3} more
                                    </div>
                                )}
                            </div>
                        </div>
                    );
                })}
            </div>
        </div>
    );
};

export default CalendarMonthView;
