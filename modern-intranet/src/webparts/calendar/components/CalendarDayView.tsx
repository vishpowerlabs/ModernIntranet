/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { format, isSameDay } from 'date-fns';
import { ICalendarEvent } from './ICalendarProps';
import styles from './Calendar.module.scss';

interface ICalendarDayViewProps {
    currentDate: Date;
    events: ICalendarEvent[];
}

const CalendarDayView: React.FC<ICalendarDayViewProps> = ({ currentDate, events }): JSX.Element => {
    const dayEvents = events.filter(ev => isSameDay(ev.start, currentDate));
    const hours = Array.from({ length: 24 }, (_, i) => i);

    return (
        <div className={styles.dayView}>
            <div className={styles.timeline}>
                {hours.map(h => (
                    <div key={h} className={styles.timeSlot}>
                        <div className={styles.timeLabel}>{format(new Date().setHours(h, 0), 'h a')}</div>
                        <div className={styles.slotContent}></div>
                    </div>
                ))}
                <div className={styles.eventsOverlay}>
                    {dayEvents.map(ev => {
                        const startHour = ev.start.getHours();
                        const startMin = ev.start.getMinutes();
                        const duration = (ev.end.getTime() - ev.start.getTime()) / (1000 * 60);
                        const top = (startHour * 60) + startMin;

                        return (
                            <div
                                key={ev.id}
                                className={styles.absoluteEvent}
                                style={{ top: `${top}px`, height: `${Math.max(duration, 30)}px` }}
                            >
                                <strong>{format(ev.start, 'h:mm a')}</strong> {ev.title}
                                {ev.location && <div className={styles.loc}>{ev.location}</div>}
                            </div>
                        );
                    })}
                </div>
            </div>
        </div>
    );
};

export default CalendarDayView;
