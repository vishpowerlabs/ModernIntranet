import { useState, useEffect, useCallback } from 'react';
import { IEmployeeDirectoryProps, IEmployee } from './IEmployeeDirectoryProps';
import { SPHttpClient, SPHttpClientResponse, MSGraphClientV3 } from '@microsoft/sp-http';

export interface IEmployeeDataState {
    employees: IEmployee[];
    loading: boolean;
    error: string | null;
    searchQuery: string;
    setSearchQuery: (query: string) => void;
    filterDept: string;
    setFilterDept: (dept: string) => void;
    filterLoc: string;
    setFilterLoc: (loc: string) => void;
    departments: string[];
    locations: string[];
    pageInfo: string;
    nextPage: () => void;
    prevPage: () => void;
    hasNextPage: boolean;
    hasPrevPage: boolean;
}

const resolveImageObject = (rawValue: string | object): any => {
    if (typeof rawValue === 'string' && rawValue.startsWith('{')) {
        try { return JSON.parse(rawValue); } catch { return null; }
    }
    return typeof rawValue === 'object' ? rawValue : null;
};

const getImageUrl = (rowItem: any, imageColumn: string, siteUrl: string): string => {
    const rawValue = rowItem[imageColumn];
    if (!rawValue) return '';

    const imageObj = resolveImageObject(rawValue as string | object);
    if (!imageObj) return typeof rawValue === 'string' ? rawValue : '';

    const url = imageObj.serverRelativeUrl || imageObj.serverUrl || imageObj.Url;
    if (url) {
        return (!url.startsWith('http') && url.startsWith('/'))
            ? `${new URL(siteUrl).origin}${url}`
            : url;
    }

    if (imageObj.fileName && rowItem.FileDirRef && rowItem.Id) {
        const origin = new URL(siteUrl).origin;
        return `${origin}${rowItem.FileDirRef}/Attachments/${rowItem.Id}/${imageObj.fileName}`;
    }

    return '';
};

// Debounce helper
function useDebounce<T>(value: T, delay: number): T {
    const [debouncedValue, setDebouncedValue] = useState<T>(value);
    useEffect(() => {
        const handler = setTimeout(() => {
            setDebouncedValue(value);
        }, delay);
        return () => clearTimeout(handler);
    }, [value, delay]);
    return debouncedValue;
}

export const useEmployeeData = (props: IEmployeeDirectoryProps): IEmployeeDataState => {
    const { context, source, siteUrl, listId, pageSize } = props;

    const [employees, setEmployees] = useState<IEmployee[]>([]);
    const [loading, setLoading] = useState<boolean>(true);
    const [error, setError] = useState<string | null>(null);

    const [searchQuery, setSearchQuery] = useState<string>('');
    const debouncedSearchQuery = useDebounce<string>(searchQuery, 300);

    const [filterDept, setFilterDept] = useState<string>('');
    const [filterLoc, setFilterLoc] = useState<string>('');

    // Pagination state
    const [pageTokens, setPageTokens] = useState<string[]>(['']); // Stack of tokens, '' is page 1
    const [currentPageIndex, setCurrentPageIndex] = useState<number>(0);
    const [nextRowToken, setNextRowToken] = useState<string | null>(null);

    // Filter dropdown options derived from currently loaded subset
    // (In a real enterprise scenario, these might be predefined terms to avoid missing options)
    const departments = Array.from(new Set(employees.map(e => e.department).filter(Boolean)));
    const locations = Array.from(new Set(employees.map(e => e.location).filter(Boolean)));

    const fetchPage = useCallback(async (token: string, query: string, dept: string, loc: string) => {
        setLoading(true);
        setError(null);
        try {
            if (source === 'graph') {
                await fetchGraphData(token, query, dept, loc);
            } else if (source === 'spList' && listId && siteUrl) {
                await fetchSPData(token, query, dept, loc);
            } else {
                setLoading(false);
            }
        } catch (err) {
            console.error("Error fetching employee data:", err);
            const errorMessage = err instanceof Error ? err.message : "An unknown error occurred.";
            setError(errorMessage);
            setLoading(false);
        }
    }, [source, siteUrl, listId, pageSize, props.nameColumn]);

    // Effect for when search/filters change (reset to page 0)
    useEffect(() => {
        setPageTokens(['']);
        setCurrentPageIndex(0);
        setNextRowToken(null);
        fetchPage('', debouncedSearchQuery, filterDept, filterLoc).catch(e => console.error(e));
    }, [debouncedSearchQuery, filterDept, filterLoc, fetchPage]);

    // Pagination handlers
    const nextPage = (): void => {
        if (nextRowToken) {
            const newTokens = [...pageTokens.slice(0, currentPageIndex + 1), nextRowToken];
            setPageTokens(newTokens);
            setCurrentPageIndex(currentPageIndex + 1);
            fetchPage(nextRowToken, debouncedSearchQuery, filterDept, filterLoc).catch(e => console.error(e));
        }
    };

    const prevPage = (): void => {
        if (currentPageIndex > 0) {
            const newIndex = currentPageIndex - 1;
            setCurrentPageIndex(newIndex);
            fetchPage(pageTokens[newIndex], debouncedSearchQuery, filterDept, filterLoc).catch(e => console.error(e));
        }
    };

    async function fetchGraphData(token: string, search: string, dept: string, loc: string): Promise<void> {
        const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient('3');

        // Note: Graph $filter requires advanced queries for many fields and consistency level eventual
        let filterStr = "accountEnabled eq true and userType eq 'Member'";

        if (search) {
            filterStr += ` and (startswith(displayName,'${search}') or startswith(jobTitle,'${search}') or startswith(department,'${search}'))`;
        }
        if (dept) {
            filterStr += ` and department eq '${dept}'`;
        }
        if (loc) {
            filterStr += ` and officeLocation eq '${loc}'`;
        }

        const requestUrl = token || `/users?$select=id,displayName,jobTitle,department,officeLocation,mail,mobilePhone,userPrincipalName&$top=${pageSize}&$filter=${filterStr}&$count=true`;

        const response = await client
            .api(requestUrl)
            .header('ConsistencyLevel', 'eventual')
            .get();

        const users = response.value || [];
        const nextUrl = response['@odata.nextLink'] || null;

        // Fetch photos in parallel, skip detailed extended profiles to avoid auto-batching $filter=id in (...) errors
        const empPromises = users.map(async (u: any) => {
            const emp: IEmployee = {
                id: u.id,
                name: u.displayName || 'Unknown',
                jobTitle: u.jobTitle || '',
                department: u.department || '',
                location: u.officeLocation || '',
                email: u.mail || u.userPrincipalName || '',
                phone: u.mobilePhone || '',
                manager: '',
                aboutMe: '',
                interests: '',
                skills: '',
                projects: ''
            };

            try {
                // Photo (blobs) - Cast responseType to any to bypass strict MSGraphClient enum check
                const photoResp = await client.api(`/users/${u.id}/photo/$value`).responseType('blob' as any).get();
                emp.photoUrl = URL.createObjectURL(photoResp);
            } catch {
                console.debug(`Failed to get graph photo value for ${u.id}`);
            }

            return emp;
        });

        const newEmps = await Promise.all(empPromises);

        setEmployees(newEmps);
        setNextRowToken(nextUrl);
        setLoading(false);
    }

    // SP List Loading
    async function fetchSPData(token: string, search: string, dept: string, loc: string): Promise<void> {
        let filterStr = '';
        const filters = [];

        if (search && props.nameColumn) {
            filters.push(`substringof('${search}',${props.nameColumn})`);
        }
        if (dept && props.departmentColumn) {
            filters.push(`(${props.departmentColumn} eq '${dept}')`);
        }
        if (loc && props.locationColumn) {
            filters.push(`(${props.locationColumn} eq '${loc}')`);
        }

        if (filters.length > 0) {
            filterStr = `&$filter=${filters.join(' and ')}`;
        }

        const selects = ['Id', 'FileDirRef', props.nameColumn, props.jobTitleColumn, props.departmentColumn, props.locationColumn, props.phoneColumn, props.photoColumn, props.projectsColumn, props.aboutMeColumn, props.interestsColumn, props.skillsColumn].filter(Boolean);
        const expands = [];

        if (props.emailColumn) {
            selects.push(`${props.emailColumn}/EMail`, `${props.emailColumn}/Title`);
            expands.push(props.emailColumn);
        }
        if (props.managerColumn) {
            selects.push(`${props.managerColumn}/Title`);
            expands.push(props.managerColumn);
        }

        const selectStr = selects.length > 0 ? `&$select=${selects.join(',')}` : '';
        const expandStr = expands.length > 0 ? `&$expand=${expands.join(',')}` : '';
        // If photo is a thumbnail column, we pull field values as text or JSON, but typically it returns a JSON string in SP REST

        // Paging logic in SP
        const skipStr = token ? `&${token}` : '';

        const url = `${siteUrl}/_api/web/lists(guid'${listId}')/items?$top=${pageSize}${selectStr}${expandStr}${filterStr}${skipStr}`;

        const response: SPHttpClientResponse = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Error ${response.status}: ${errorText}`);
        }

        const json = await response.json();
        const items = json.value || [];

        // Parse SP next page token
        let nextTokenStr = null;
        if (json['@odata.nextLink']) {
            // Extract skiptoken from nextLink
            const nextLinkObj = new URL(json['@odata.nextLink']);
            nextTokenStr = nextLinkObj.search.substring(1).split('&').find(p => p.startsWith('%24skiptoken=') || p.startsWith('$skiptoken='));
        }

        const newEmps = items.map((item: any) => {
            let emailVal = '';
            if (props.emailColumn && item[props.emailColumn]) {
                emailVal = item[props.emailColumn].EMail || item[props.emailColumn].Title || '';
            }

            let mgrVal = '';
            if (props.managerColumn && item[props.managerColumn]) {
                mgrVal = item[props.managerColumn].Title || '';
            }

            const photoUrl = props.photoColumn ? getImageUrl(item, props.photoColumn, siteUrl) : '';

            return {
                id: item.Id.toString(),
                name: props.nameColumn ? (item[props.nameColumn] || '') : '',
                jobTitle: props.jobTitleColumn ? (item[props.jobTitleColumn] || '') : '',
                department: props.departmentColumn ? (item[props.departmentColumn] || '') : '',
                location: props.locationColumn ? (item[props.locationColumn] || '') : '',
                phone: props.phoneColumn ? (item[props.phoneColumn] || '') : '',
                email: emailVal,
                manager: mgrVal,
                photoUrl: photoUrl,
                projects: props.projectsColumn ? (item[props.projectsColumn] || '') : '',
                aboutMe: props.aboutMeColumn ? (item[props.aboutMeColumn] || '') : '',
                interests: props.interestsColumn ? (item[props.interestsColumn] || '') : '',
                skills: props.skillsColumn ? (item[props.skillsColumn] || '') : ''
            } as IEmployee;
        });

        setEmployees(newEmps);
        setNextRowToken(nextTokenStr || null);
        setLoading(false);
    }

    const pageInfo = `Page ${currentPageIndex + 1}`;

    return {
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
        hasNextPage: !!nextRowToken,
        hasPrevPage: currentPageIndex > 0
    };
};
