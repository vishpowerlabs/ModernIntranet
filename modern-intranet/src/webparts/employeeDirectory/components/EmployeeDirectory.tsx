import * as React from 'react';
import styles from './EmployeeDirectory.module.scss';
import { IEmployeeDirectoryProps, IEmployee } from './IEmployeeDirectoryProps';
import { useEmployeeData } from './useEmployeeData';
import EmployeeListView from './EmployeeListView';
import EmployeeGridView from './EmployeeGridView';
import PaginationBar from './PaginationBar';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { EmptyState } from '../../../common/components/EmptyState/EmptyState';
import { WebPartHeader } from '../../../common/components/WebPartHeader/WebPartHeader';

const EmployeeDirectory: React.FC<IEmployeeDirectoryProps> = (props): JSX.Element => {
    const {
        showTitle,
        title,
        showBackgroundBar,
        viewMode: defaultViewMode,
        titleBarStyle,
        showFilters,
        showPagination
    } = props;

    const [viewMode, setViewMode] = React.useState<'list' | 'grid'>(defaultViewMode);
    const [selectedEmployee, setSelectedEmployee] = React.useState<IEmployee | null>(null);
    const [isPanelLoading, setIsPanelLoading] = React.useState<boolean>(false);

    // Sync viewMode if prop changes
    React.useEffect(() => {
        setViewMode(defaultViewMode);
    }, [defaultViewMode]);

    // Lazy load graph user details when panel is open
    React.useEffect(() => {
        if (!selectedEmployee || props.source !== 'graph' || selectedEmployee.manager || selectedEmployee.aboutMe || selectedEmployee.skills || selectedEmployee.interests || selectedEmployee.projects) {
            return; // Only load if we have a graph user that lacks these details
        }

        const fetchDetails = async (): Promise<void> => {
            setIsPanelLoading(true);
            try {
                const client = await props.context.msGraphClientFactory.getClient('3');

                // Details
                const profileResp = await client.api(`/users/${selectedEmployee.id}?$select=aboutMe,interests,skills,pastProjects`).get();
                const newDetails: Partial<IEmployee> = {};
                if (profileResp.aboutMe) newDetails.aboutMe = profileResp.aboutMe;
                if (profileResp.interests) newDetails.interests = profileResp.interests.join(', ');
                if (profileResp.skills) newDetails.skills = profileResp.skills.join(', ');
                if (profileResp.pastProjects) newDetails.projects = profileResp.pastProjects.join(', ');

                // Manager
                try {
                    const mgrResp = await client.api(`/users/${selectedEmployee.id}/manager?$select=displayName`).get();
                    if (mgrResp?.displayName) newDetails.manager = mgrResp.displayName;
                } catch (mgrError) {
                    console.debug('Failed to fetch manager details', mgrError);
                }

                if (Object.keys(newDetails).length > 0) {
                    setSelectedEmployee(prev => prev ? { ...prev, ...newDetails } : prev);
                }
            } catch (err) {
                console.error("Error fetching lazy loaded graph details:", err);
            } finally {
                setIsPanelLoading(false);
            }
        };

        fetchDetails().catch(e => console.error(e));
    }, [selectedEmployee?.id, props.source, props.context]);

    const {
        employees,
        loading,
        error,
        searchQuery,
        setSearchQuery,
        filterDept,
        setFilterDept,
        filterLoc,
        setFilterLoc,
        departments,
        locations,
        pageInfo,
        nextPage,
        prevPage,
        hasNextPage,
        hasPrevPage
    } = useEmployeeData(props);

    const handleSearchChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
        setSearchQuery(e.target.value);
    };

    const renderHeader = (): JSX.Element => (
        <WebPartHeader
            title={title || ''}
            showTitle={!!showTitle}
            showBackgroundBar={!!showBackgroundBar}
            titleBarStyle={titleBarStyle || 'underline'}
        />
    );

    if (props.source === 'spList' && (!props.siteUrl || !props.listId || !props.nameColumn)) {
        return (
            <div className={styles.employeeDirectory}>
                <EmptyState
                    icon="Contact"
                    title="Employee Directory - Configuration Required"
                    message="Please complete the web part configuration to display employees."
                    description="You have selected SharePoint List as the source. You need to specify the Site URL, List ID, and map the required columns (Name, Photo, etc.) in the property pane."
                />
            </div>
        );
    }

    if (error) {
        return (
            <div className={styles.employeeDirectory}>
                <EmptyState
                    icon="Error"
                    title="Error Loading Directory"
                    message={error}
                    description="Please verify your data source settings and user permissions."
                />
            </div>
        );
    }

    if (loading) {
        return (
            <div className={styles.employeeDirectory}>
                {renderHeader()}
                <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '300px' }}>
                    <Spinner size={SpinnerSize.large} label="Loading employees..." />
                </div>
            </div>
        );
    }

    if (employees.length === 0) {
        return (
            <div className={styles.employeeDirectory}>
                <EmptyState
                    icon="People"
                    title="No Employees Found"
                    message={searchQuery || filterDept || filterLoc ? "No results match your search or filter criteria." : "There are no employees to display."}
                    description="Try adjusting your search terms or verify that the data source contains employee records."
                />
            </div>
        );
    }

    return (
        <div className={styles.employeeDirectory}>
            {renderHeader()}

            <div className={styles.toolbar}>
                <div className={styles.search}>
                    <svg viewBox="0 0 24 24">
                        <path d="M15.5 14h-.79l-.28-.27A6.471 6.471 0 0016 9.5 6.5 6.5 0 109.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z" />
                    </svg>
                    <input
                        type="text"
                        placeholder="Search by name, department, or title…"
                        value={searchQuery}
                        onChange={handleSearchChange}
                    />
                </div>
                <div className={styles.viewToggle}>
                    <button
                        className={`${styles.viewBtn} ${viewMode === 'list' ? styles.active : ''}`}
                        onClick={() => setViewMode('list')}
                        title="List View"
                    >
                        ☰
                    </button>
                    <button
                        className={`${styles.viewBtn} ${viewMode === 'grid' ? styles.active : ''}`}
                        onClick={() => setViewMode('grid')}
                        title="Grid View"
                    >
                        ⊞
                    </button>
                </div>
            </div>

            {showFilters && (
                <div className={styles.filters}>
                    <select value={filterDept} onChange={(e) => setFilterDept(e.target.value)}>
                        <option value="">All Departments</option>
                        {departments.map(d => <option key={d} value={d}>{d}</option>)}
                    </select>
                    <select value={filterLoc} onChange={(e) => setFilterLoc(e.target.value)}>
                        <option value="">All Locations</option>
                        {locations.map(l => <option key={l} value={l}>{l}</option>)}
                    </select>
                </div>
            )}

            {viewMode === 'list' ? (
                <EmployeeListView employees={employees} onEmployeeClick={setSelectedEmployee} />
            ) : (
                <EmployeeGridView employees={employees} onEmployeeClick={setSelectedEmployee} />
            )}

            {showPagination && (
                <PaginationBar
                    pageInfo={pageInfo}
                    onPrev={prevPage}
                    onNext={nextPage}
                    hasPrev={hasPrevPage}
                    hasNext={hasNextPage}
                />
            )}

            <Panel
                isOpen={!!selectedEmployee}
                onDismiss={() => setSelectedEmployee(null)}
                isLightDismiss={true}
                type={PanelType.medium}
                headerText="Employee Profile"
                closeButtonAriaLabel="Close"
            >
                {selectedEmployee && (
                    <div className={styles.profilePanel}>
                        <div className={styles.profileHeader}>
                            {selectedEmployee.photoUrl ? (
                                <img src={selectedEmployee.photoUrl} alt={selectedEmployee.name} className={styles.profilePhoto} />
                            ) : (
                                <div className={styles.profileInitials}>
                                    {selectedEmployee.name.split(' ').map(n => n[0]).join('').substring(0, 2)}
                                </div>
                            )}
                            <div className={styles.profileTitles}>
                                <h2>{selectedEmployee.name}</h2>
                                <p className={styles.jobTitle}>{selectedEmployee.jobTitle}</p>
                                <p className={styles.department}>{selectedEmployee.department}</p>
                            </div>
                        </div>

                        <div className={styles.profileContact}>
                            <h3>Contact Information</h3>
                            {selectedEmployee.email && (
                                <div className={styles.contactRow}>
                                    <span className={styles.icon}>✉</span>
                                    <a href={`mailto:${selectedEmployee.email}`}>{selectedEmployee.email}</a>
                                </div>
                            )}
                            {selectedEmployee.phone && (
                                <div className={styles.contactRow}>
                                    <span className={styles.icon}>📞</span>
                                    <a href={`tel:${selectedEmployee.phone}`}>{selectedEmployee.phone}</a>
                                </div>
                            )}
                            {selectedEmployee.location && (
                                <div className={styles.contactRow}>
                                    <span className={styles.icon}>🏢</span>
                                    <span>{selectedEmployee.location}</span>
                                </div>
                            )}
                        </div>

                        {selectedEmployee.manager && (
                            <div className={styles.profileOrganization}>
                                <h3>Organization</h3>
                                <div className={styles.contactRow}>
                                    <span className={styles.icon}>👤</span>
                                    <span><strong>Manager:</strong> {selectedEmployee.manager}</span>
                                </div>
                            </div>
                        )}

                        {isPanelLoading ? (
                            <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '100px', marginTop: '20px' }}>
                                <Spinner size={SpinnerSize.medium} label="Loading details..." />
                            </div>
                        ) : (
                            (selectedEmployee.projects || selectedEmployee.aboutMe || selectedEmployee.interests || selectedEmployee.skills) && (
                                <div className={styles.profileOrganization}>
                                    <h3>Additional Details</h3>
                                    {selectedEmployee.aboutMe && (
                                        <div className={styles.detailSection}>
                                            <strong>About Me:</strong>
                                            <div dangerouslySetInnerHTML={{ __html: selectedEmployee.aboutMe.split('\n').join('<br/>') }} />
                                        </div>
                                    )}
                                    {selectedEmployee.projects && (
                                        <div className={styles.detailSection}>
                                            <strong>Projects:</strong>
                                            <div dangerouslySetInnerHTML={{ __html: selectedEmployee.projects.split('\n').join('<br/>') }} />
                                        </div>
                                    )}
                                    {selectedEmployee.skills && (
                                        <div className={styles.detailSection}>
                                            <strong>Skills:</strong>
                                            <div dangerouslySetInnerHTML={{ __html: selectedEmployee.skills.split('\n').join('<br/>') }} />
                                        </div>
                                    )}
                                    {selectedEmployee.interests && (
                                        <div className={styles.detailSection}>
                                            <strong>Interests:</strong>
                                            <div dangerouslySetInnerHTML={{ __html: selectedEmployee.interests.split('\n').join('<br/>') }} />
                                        </div>
                                    )}
                                </div>
                            )
                        )}
                    </div>
                )}
            </Panel>
        </div>
    );
};

export default EmployeeDirectory;
