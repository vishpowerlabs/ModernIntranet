/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { IconButton, DefaultButton } from '@fluentui/react';
import { format } from 'date-fns';
import { CalendarViewType } from './useCalendarNavigation';
import styles from './Calendar.module.scss';

interface ICalendarHeaderProps {
    currentDate: Date;
    view: CalendarViewType;
    onNavigate: (direction: 'next' | 'prev') => void;
    onToday: () => void;
    onViewChange: (view: CalendarViewType) => void;
}

const CalendarHeader: React.FC<ICalendarHeaderProps> = (props) => {
    const { currentDate, view, onNavigate, onToday, onViewChange } = props;

    const getContextLabel = () => {
        switch (view) {
            case 'day': return format(currentDate, 'MMMM d, yyyy');
            case 'week': {
                const start = format(currentDate, 'MMM d');
                const end = format(new Date(currentDate.getTime() + 6 * 24 * 60 * 60 * 1000), 'd, yyyy');
                return `${start} – ${end}`;
            }
            case 'month': return format(currentDate, 'MMMM yyyy');
            case 'year': return format(currentDate, 'yyyy');
        }
    };

    return (
        <div className={styles.header}>
            <div className={styles.headerLeft}>
                <IconButton iconProps={{ iconName: 'ChevronLeft' }} onClick={() => onNavigate('prev')} />
                <IconButton iconProps={{ iconName: 'ChevronRight' }} onClick={() => onNavigate('next')} />
                <span className={styles.contextLabel}>{getContextLabel()}</span>
            </div>

            <div className={styles.headerCenter}>
                <DefaultButton text="Today" onClick={onToday} />
            </div>

            <div className={styles.headerRight}>
                <div className={styles.viewSegmentedButton}>
                    {(['day', 'week', 'month', 'year'] as CalendarViewType[]).map(v => (
                        <button
                            key={v}
                            className={`${styles.viewBtn} ${view === v ? styles.active : ''}`}
                            onClick={() => onViewChange(v)}
                        >
                            {v.charAt(0).toUpperCase() + v.slice(1)}
                        </button>
                    ))}
                </div>
            </div>
        </div>
    );
};

export default CalendarHeader;
