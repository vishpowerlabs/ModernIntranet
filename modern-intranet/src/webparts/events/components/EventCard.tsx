import * as React from 'react';
import styles from './Events.module.scss';
import { Icon } from '@fluentui/react/lib/Icon';

export interface IEventCardProps {
  title: string;
  date: Date;
  imageUrl?: string;
  location?: string;
  linkUrl?: string;
}

export const EventCard: React.FC<IEventCardProps> = ({ title, date, imageUrl, location, linkUrl }) => {
  const day = date.getDate();
  const month = date.toLocaleString('default', { month: 'short' });
  const time = date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

  // Fallback icon name if image missing
  const cardStyle = imageUrl ? { backgroundImage: `url('${imageUrl}')` } : {};

  return (
    <a 
      href={linkUrl || '#'} 
      className={styles.eventCard}
      target={linkUrl ? "_blank" : "_self"}
      rel="noopener noreferrer"
    >
      <div className={styles.imageArea} style={cardStyle}>
        <div className={styles.dateBadge}>
          <span className={styles.day}>{day}</span>
          <span className={styles.month}>{month}</span>
        </div>
      </div>
      <div className={styles.contentArea}>
        <h3 className={styles.title}>{title}</h3>
        <div className={styles.metaInfo}>
          <div className={styles.infoItem}>
            <Icon iconName="Clock" className={styles.icon} />
            <span>{time}</span>
          </div>
          {location && (
            <div className={styles.infoItem}>
              <Icon iconName="MapPin" className={styles.icon} />
              <span>{location}</span>
            </div>
          )}
        </div>
      </div>
    </a>
  );
};
