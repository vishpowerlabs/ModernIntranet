/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { Nav, INavLinkGroup } from '@fluentui/react/lib/Nav';

export interface ISideNavProps {
    categories: string[];
    selectedCategory: string;
    onSelectCategory: (category: string) => void;
}

export const SideNav: React.FunctionComponent<ISideNavProps> = (props) => {
    const { categories, selectedCategory, onSelectCategory } = props;

    const navGroups: INavLinkGroup[] = [
        {
            links: categories.map((cat) => ({
                name: cat,
                url: '',
                key: cat,
                isExpanded: true,
                onClick: (ev, item) => {
                    if (ev) ev.preventDefault(); // Prevent navigation
                    if (item?.key) onSelectCategory(item.key);
                }
            })),
        },
    ];

    return (
        <Nav
            groups={navGroups}
            selectedKey={selectedCategory}
            styles={{
                root: {
                    width: '100%',
                    boxSizing: 'border-box',
                    overflowY: 'auto'
                }
            }}
        />
    );
};
