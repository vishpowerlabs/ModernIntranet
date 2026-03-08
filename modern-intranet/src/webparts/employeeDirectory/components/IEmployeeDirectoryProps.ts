import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IEmployeeDirectoryProps {
    context: WebPartContext;
    showTitle: boolean;
    title: string;
    showBackgroundBar: boolean;
    titleBarStyle: 'solid' | 'underline';

    source: 'graph' | 'spList';
    siteUrl: string;
    listId: string;

    nameColumn: string;
    photoColumn: string;
    jobTitleColumn: string;
    departmentColumn: string;
    locationColumn: string;
    emailColumn: string;
    phoneColumn: string;
    managerColumn: string;
    projectsColumn: string;
    aboutMeColumn: string;
    interestsColumn: string;
    skillsColumn: string;

    viewMode: 'list' | 'grid';
    pageSize: number;
    showFilters: boolean;
    showPagination: boolean;
}

export interface IEmployee {
    id: string;
    name: string;
    photoUrl?: string;
    jobTitle: string;
    department: string;
    location: string;
    email: string;
    phone: string;
    manager: string;
    projects?: string;
    aboutMe?: string;
    interests?: string;
    skills?: string;
}
