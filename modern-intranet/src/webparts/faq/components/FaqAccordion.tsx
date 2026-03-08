import * as React from 'react';
import styles from './Faq.module.scss';
import { IFaqItem } from './IFaqProps';
import FaqItem from './FaqItem';

interface IFaqAccordionProps {
    items: IFaqItem[];
    expandedIds: Set<number>;
    toggleItem: (id: number) => void;
    searchQuery: string;
}

const FaqAccordion: React.FC<IFaqAccordionProps> = ({ items, expandedIds, toggleItem, searchQuery }) => {
    return (
        <section className={styles.accordion} aria-label="FAQ Accordion">
            {items.map(item => (
                <FaqItem
                    key={item.id}
                    item={item}
                    isExpanded={expandedIds.has(item.id)}
                    onToggle={() => toggleItem(item.id)}
                    searchQuery={searchQuery}
                />
            ))}
        </section>
    );
};

export default FaqAccordion;
