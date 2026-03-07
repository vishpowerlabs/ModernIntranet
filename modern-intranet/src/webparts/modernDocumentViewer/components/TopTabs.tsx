/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import * as React from 'react';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';

export interface ITopTabsProps {
    subCategories: string[];
    selectedSubCategory: string;
    onSelectSubCategory: (subCat: string) => void;
}

export const TopTabs: React.FunctionComponent<ITopTabsProps> = (props) => {
    const { subCategories, selectedSubCategory, onSelectSubCategory } = props;

    const handleLinkClick = (item?: PivotItem): void => {
        if (item) {
            onSelectSubCategory(item.props.itemKey || '');
        }
    };

    return (
        <Pivot
            selectedKey={selectedSubCategory}
            onLinkClick={handleLinkClick}
            headersOnly={true}
        >
            <PivotItem headerText="All" itemKey="" />
            {subCategories.filter(s => s !== 'All').map((sub) => (
                <PivotItem headerText={sub} itemKey={sub} key={sub} />
            ))}
        </Pivot>
    );
};
