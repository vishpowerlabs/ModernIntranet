import * as React from 'react';
import styles from './EmployeeDirectory.module.scss';

export interface IPaginationBarProps {
    pageInfo: string;
    onPrev: () => void;
    onNext: () => void;
    hasPrev: boolean;
    hasNext: boolean;
}

const PaginationBar: React.FC<IPaginationBarProps> = ({ pageInfo, onPrev, onNext, hasPrev, hasNext }) => {
    return (
        <div className={styles.paginationBar}>
            <span className={styles.pageInfo}>{pageInfo}</span>
            <div className={styles.pageBtns}>
                <button
                    className={styles.pageBtn}
                    onClick={onPrev}
                    disabled={!hasPrev}
                >
                    &lsaquo; Previous
                </button>
                <button
                    className={styles.pageBtn}
                    onClick={onNext}
                    disabled={!hasNext}
                >
                    Next &rsaquo;
                </button>
            </div>
        </div>
    );
};

export default PaginationBar;
