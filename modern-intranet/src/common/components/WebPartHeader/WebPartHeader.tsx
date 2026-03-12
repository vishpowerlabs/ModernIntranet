/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import styles from './WebPartHeader.module.scss';

export interface IWebPartHeaderProps {
  title: string;
  showTitle: boolean;
  showBackgroundBar: boolean;
  titleBarStyle: 'solid' | 'underline';
  description?: string;
  descriptionFontSize?: string;
  onActionClick?: () => void;
  actionTitle?: string;
  actionUrl?: string;
}

export const WebPartHeader: React.FC<IWebPartHeaderProps> = (props) => {
  if (!props.showTitle || !props.title) return null;

  const getHeaderClasses = (): string => {
    let classes = styles.webpartHeader;
    if (props.showBackgroundBar) {
      if (props.titleBarStyle === 'solid') {
        classes += ` ${styles.solidBackground}`;
      } else {
        classes += ` ${styles.underlineBackground}`;
      }
    }
    return classes;
  };

  return (
    <div className={getHeaderClasses()}>
      <div className={styles.titleContainer}>
        <h2>{props.title}</h2>
      </div>
      {props.actionTitle && (
        <div className={styles.actionContainer}>
          {props.actionUrl ? (
            <a href={props.actionUrl} target="_blank" rel="noopener noreferrer">
              {props.actionTitle}
            </a>
          ) : (
            <button onClick={props.onActionClick} type="button">
              {props.actionTitle}
            </button>
          )}
        </div>
      )}
      {props.description && (
        <div 
          className={styles.descriptionContainer}
          style={{ fontSize: props.descriptionFontSize || '14px' }}
        >
          {props.description}
        </div>
      )}
    </div>
  );
};
