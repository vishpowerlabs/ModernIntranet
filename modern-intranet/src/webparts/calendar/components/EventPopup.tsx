/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { format } from 'date-fns';
import { ICalendarEvent } from './ICalendarProps';
import { Icon } from '@fluentui/react';
import styles from './Calendar.module.scss';

interface IEventPopupProps {
    event: ICalendarEvent;
    position: { top: number; left: number };
    onClose: () => void;
}

const EventPopup: React.FC<IEventPopupProps> = ({ event, position, onClose }) => {
    return (
        <div
            className={styles.eventPopup}
            style={{ top: position.top, left: position.left }}
        >
            <div className={styles.popupHeader}>
                <div className={styles.popupTitle}>{event.title}</div>
                <Icon iconName="Cancel" className={styles.closeIcon} onClick={onClose} />
            </div>
            <div className={styles.popupTime}>
                <Icon iconName="Clock" />
                {format(event.start, 'MMM d, yyyy h:mm a')}
                {event.end && ` - ${format(event.end, 'h:mm a')}`}
            </div>
            {event.location && (
                <div className={styles.popupLocation}>
                    <Icon iconName="Poi" />
                    {event.location}
                </div>
            )}
        </div>
    );
};

export default EventPopup;
