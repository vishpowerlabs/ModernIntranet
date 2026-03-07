/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { format, startOfWeek, eachDayOfInterval, isSameDay, isToday } from 'date-fns';
import { ICalendarEvent } from './ICalendarProps';
import styles from './Calendar.module.scss';

interface ICalendarWeekViewProps {
    currentDate: Date;
    events: ICalendarEvent[];
}

const CalendarWeekView: React.FC<ICalendarWeekViewProps> = ({ currentDate, events }): JSX.Element => {
    const startDate = startOfWeek(currentDate);
    const days = eachDayOfInterval({ start: startDate, end: new Date(startDate.getTime() + 6 * 24 * 60 * 60 * 1000) });

    const getEventsForDay = (day: Date): ICalendarEvent[] => {
        return events.filter(ev => isSameDay(ev.start, day));
    };

    return (
        <div className={styles.weekView}>
            <div className={styles.weekGrid}>
                {days.map((day, i) => (
                    <div key={day.toISOString()} className={styles.weekDayCol}>
                        <div className={`${styles.dayHeader} ${isToday(day) ? styles.today : ''}`}>
                            <span className={styles.dayName}>{format(day, 'EEE')}</span>
                            <span className={styles.dayNum}>{format(day, 'd')}</span>
                        </div>
                        <div className={styles.dayEvents}>
                            {getEventsForDay(day).map(ev => (
                                <div key={ev.id} className={styles.eventCard}>
                                    <div className={styles.eventTime}>{format(ev.start, 'h:mm a')}</div>
                                    <div className={styles.eventTitle}>{ev.title}</div>
                                </div>
                            ))}
                        </div>
                    </div>
                ))}
            </div>
        </div>
    );
};

export default CalendarWeekView;
