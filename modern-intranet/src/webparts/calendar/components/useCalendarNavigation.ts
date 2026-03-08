/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import { useState, useCallback } from 'react';
import {
    addDays,
    addWeeks,
    addMonths,
    addYears,
    startOfDay,
    startOfWeek,
    startOfMonth,
    startOfYear,
    endOfDay,
    endOfWeek,
    endOfMonth,
    endOfYear
} from 'date-fns';

export type CalendarViewType = 'day' | 'week' | 'month' | 'year';

export interface ICalendarNavigation {
    currentDate: Date;
    view: CalendarViewType;
    navigate: (direction: 'next' | 'prev') => void;
    snapToToday: () => void;
    changeView: (newView: CalendarViewType) => void;
    setSpecificDate: (date: Date, newView?: CalendarViewType) => void;
    getViewRange: () => { start: Date; end: Date; } | undefined;
}

export const useCalendarNavigation = (defaultView: CalendarViewType): ICalendarNavigation => {
    const [currentDate, setCurrentDate] = useState(new Date());
    const [view, setView] = useState<CalendarViewType>(defaultView);

    const navigate = useCallback((direction: 'next' | 'prev') => {
        const amount = direction === 'next' ? 1 : -1;
        switch (view) {
            case 'day': return setCurrentDate(prev => addDays(prev, amount));
            case 'week': return setCurrentDate(prev => addWeeks(prev, amount));
            case 'month': return setCurrentDate(prev => addMonths(prev, amount));
            case 'year': return setCurrentDate(prev => addYears(prev, amount));
        }
    }, [view]);

    const snapToToday = useCallback(() => {
        setCurrentDate(new Date());
    }, []);

    const changeView = useCallback((newView: CalendarViewType) => {
        setView(newView);
    }, []);

    const setSpecificDate = useCallback((date: Date, newView?: CalendarViewType) => {
        setCurrentDate(date);
        if (newView) setView(newView);
    }, []);

    const getViewRange = useCallback(() => {
        switch (view) {
            case 'day': return { start: startOfDay(currentDate), end: endOfDay(currentDate) };
            case 'week': return { start: startOfWeek(currentDate), end: endOfWeek(currentDate) };
            case 'month': return { start: startOfMonth(currentDate), end: endOfMonth(currentDate) };
            case 'year': return { start: startOfYear(currentDate), end: endOfYear(currentDate) };
        }
    }, [currentDate, view]);

    return {
        currentDate,
        view,
        navigate,
        snapToToday,
        changeView,
        setSpecificDate,
        getViewRange
    };
};
