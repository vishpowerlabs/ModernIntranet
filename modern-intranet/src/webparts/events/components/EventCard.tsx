/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import styles from './Events.module.scss';
import { Icon } from '@fluentui/react/lib/Icon';

export interface IEventCardProps {
  id: string;
  title: string;
  date: Date;
  location?: string;
  description?: string;
  linkUrl?: string;
  layout: 'list' | 'grid' | 'compact';
  onInfoClick?: (event: React.MouseEvent<HTMLElement>, id: string) => void;
}

export const EventCard: React.FC<IEventCardProps> = ({ 
  id, 
  title, 
  date, 
  location, 
  description, 
  linkUrl, 
  layout, 
  onInfoClick 
}) => {
  const day = date.getDate();
  const month = date.toLocaleString('default', { month: 'short' });
  const year = date.getFullYear();

  const handleInfoClick = (e: React.MouseEvent<HTMLElement>): void => {
    e.preventDefault();
    e.stopPropagation();
    if (onInfoClick) {
      onInfoClick(e, id);
    }
  };

  const pinSVG = <Icon iconName="MapPin" className={styles.icon} />;
  const infoSVG = <Icon iconName="Info" />;

  if (layout === 'list') {
    return (
      <div className={styles.eventListItem} onClick={(e) => onInfoClick?.(e, id)}>
        <div className={styles.dateBadge}>
          <span className={styles.month}>{month}</span>
          <span className={styles.day}>{day}</span>
        </div>
        <div className={styles.listInfo}>
          <div className={styles.eventListTitle}>{title}</div>
          <div className={styles.eventListLoc}>{pinSVG} {location || 'No location'}</div>
        </div>
        <button className={styles.infoBtn} onClick={handleInfoClick}>{infoSVG}</button>
      </div>
    );
  }

  if (layout === 'grid') {
    return (
      <div className={styles.eventCard} onClick={(e) => onInfoClick?.(e, id)}>
        <div className={styles.cardTop}>
          <button className={styles.cardInfoBtn} onClick={handleInfoClick}>{infoSVG}</button>
          <div className={styles.monthLbl}>{month} {year}</div>
          <div className={styles.dayLbl}>{day}</div>
        </div>
        <div className={styles.cardBody}>
          <div className={styles.cardTitle}>{title}</div>
          <div className={styles.cardLoc}>{pinSVG} {location || 'No location'}</div>
          {description && <div className={styles.cardDesc}>{description}</div>}
        </div>
      </div>
    );
  }

  // Compact Layout
  return (
    <div className={styles.eventCompactItem} onClick={(e) => onInfoClick?.(e, id)}>
      <div className={styles.compactPill}>{month} {day}</div>
      <div className={styles.compactInfo}>
        <div className={styles.compactTitle}>{title}</div>
        <div className={styles.compactLoc}>{location || 'No location'}</div>
      </div>
      <button className={styles.infoBtn} onClick={handleInfoClick}>{infoSVG}</button>
    </div>
  );
};
