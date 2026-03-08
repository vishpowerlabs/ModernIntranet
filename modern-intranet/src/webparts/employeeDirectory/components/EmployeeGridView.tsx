import * as React from 'react';
import styles from './EmployeeDirectory.module.scss';
import { IEmployee } from './IEmployeeDirectoryProps';

export interface IEmployeeGridViewProps {
    employees: IEmployee[];
    onEmployeeClick: (employee: IEmployee) => void;
}

const EmployeeCard: React.FC<{ employee: IEmployee, onClick: () => void }> = ({ employee, onClick }) => {
    const renderPhotoArea = (): JSX.Element => {
        if (employee.photoUrl) {
            return (
                <div className={styles.cardPhotoArea}>
                    <img src={employee.photoUrl} alt={employee.name} />
                </div>
            );
        }
        const initials = employee.name.split(' ').map(n => n[0]).join('').substring(0, 2);
        const colors = ['#0078d4', '#038387', '#8764b8', '#c239b3', '#ca5010', '#498205'];
        const codeValue = employee.name.codePointAt(0);
        const color = colors[(codeValue || 0) % colors.length];

        return (
            <div className={styles.cardPhotoArea}>
                <div className={styles.avatarBase} style={{ background: color, width: 64, height: 64, fontSize: 22 }}>
                    {initials}
                </div>
            </div>
        );
    };

    return (
        <div
            className={styles.gridCard}
            title={employee.location ? `Office: ${employee.location}` : undefined}
            onClick={onClick}
            role="button"
            tabIndex={0}
            onKeyDown={(e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    onClick();
                }
            }}
        >
            {renderPhotoArea()}
            <div className={styles.cardBody}>
                <div className={styles.cardName}>{employee.name}</div>
                <div className={styles.cardRole}>{employee.jobTitle} &middot; {employee.department}</div>

                {employee.email && (
                    <div className={styles.emailContainerCentered}>
                        <a
                            href={`mailto:${employee.email}`}
                            className={styles.emailText}
                            onClick={(e) => e.stopPropagation()}
                            title="Send Email"
                        >
                            {employee.email}
                        </a>
                        <button
                            className={styles.copyBtn}
                            title="Copy Email"
                            onClick={(e) => {
                                e.stopPropagation();
                                navigator.clipboard.writeText(employee.email).catch(console.error);
                            }}
                        >
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">
                                <path d="M19 21H8V7h11m0-2H8a2 2 0 00-2 2v14a2 2 0 002 2h11a2 2 0 002-2V7a2 2 0 00-2-2m-3-4H4a2 2 0 00-2 2v14h2V3h12V1z" />
                            </svg>
                        </button>
                    </div>
                )}

                <div className={styles.cardActions}>
                    <a
                        className={styles.actionBtn}
                        title="Send Email"
                        href={`mailto:${employee.email}`}
                        onClick={(e) => e.stopPropagation()}
                    >
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">
                            <path d="M22 6c0-1.1-.9-2-2-2H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6zm-2 0l-8 5-8-5h16zm0 12H4V8l8 5 8-5v10z" />
                        </svg>
                    </a>
                    {employee.email && (
                        <a
                            className={styles.actionBtn}
                            title="Teams Call"
                            href={`msteams://teams.microsoft.com/l/call/0/0?users=${employee.email}`}
                            onClick={(e) => e.stopPropagation()}
                        >
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">
                                <path d="M20 15.5c-1.25 0-2.45-.2-3.57-.57a1.02 1.02 0 00-1.02.24l-2.2 2.2a15.045 15.045 0 01-6.59-6.59l2.2-2.21a.96.96 0 00.25-1A11.36 11.36 0 013.5 4 1 1 0 002.5 5C2.5 14.66 10.34 22.5 20 22.5a1 1 0 001-1v-3.5a1 1 0 00-1-1z" />
                            </svg>
                        </a>
                    )}
                </div>
            </div>
        </div>
    );
};

const EmployeeGridView: React.FC<IEmployeeGridViewProps> = ({ employees, onEmployeeClick }) => {
    return (
        <div className={styles.gridView}>
            {employees.map(emp => (
                <EmployeeCard key={emp.id} employee={emp} onClick={() => onEmployeeClick(emp)} />
            ))}
        </div>
    );
};

export default EmployeeGridView;
