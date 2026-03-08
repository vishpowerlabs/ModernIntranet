import * as React from 'react';
import styles from './Faq.module.scss';
import { IFaqItem } from './IFaqProps';

interface IFaqItemProps {
    item: IFaqItem;
    isExpanded: boolean;
    onToggle: () => void;
    searchQuery: string;
}

const FaqItem: React.FC<IFaqItemProps> = ({ item, isExpanded, onToggle, searchQuery }) => {

    const highlightText = (text: string, query: string): React.ReactNode => {
        if (!query || !text) return text;

        const lowerText = text.toLowerCase();
        const lowerQuery = query.toLowerCase();
        const parts: React.ReactNode[] = [];
        let startIndex = 0;

        while (startIndex < text.length) {
            const index = lowerText.indexOf(lowerQuery, startIndex);
            if (index === -1) {
                parts.push(text.substring(startIndex));
                break;
            }

            if (index > startIndex) {
                parts.push(text.substring(startIndex, index));
            }
            parts.push(<mark key={index}>{text.substring(index, index + query.length)}</mark>);
            startIndex = index + query.length;
        }

        return <span>{parts}</span>;
    };

    return (
        <div className={`${styles.faqItem} ${isExpanded ? styles.expanded : ''}`}>
            <button
                className={styles.itemHeader}
                onClick={onToggle}
                aria-expanded={isExpanded}
                aria-controls={`faq-answer-${item.id}`}
                id={`faq-question-${item.id}`}
            >
                <div className={styles.questionText}>
                    {highlightText(item.question, searchQuery)}
                </div>
                <svg className={styles.chevron} viewBox="0 0 24 24">
                    <path d="M10 6L8.59 7.41 13.17 12l-4.58 4.59L10 18l6-6z" />
                </svg>
            </button>
            <section
                className={styles.itemBody}
                id={`faq-answer-${item.id}`}
                aria-labelledby={`faq-question-${item.id}`}
            >
                <div className={styles.answerText}>
                    <div dangerouslySetInnerHTML={{ __html: item.answer }} />
                    {item.category && (
                        <span className={styles.categoryBadge}>{item.category}</span>
                    )}
                </div>
            </section>
        </div>
    );
};

export default FaqItem;
