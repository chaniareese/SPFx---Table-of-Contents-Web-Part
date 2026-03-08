import * as React from 'react';
import styles from './TableOfContents.module.scss';
import { IHeading } from './ITableOfContentsState';

export interface ITOCNavigatorProps {
  headings: IHeading[];
  activeHeadingId: string | null;
  showH2: boolean;
  showH3: boolean;
  showH4: boolean;
  tocTitle?: string;
  onHeadingClick: (id: string) => void;
}

export const TOCNavigator: React.FC<ITOCNavigatorProps> = (props) => {
  const { headings, activeHeadingId, showH2, showH3, showH4, tocTitle, onHeadingClick } = props;

  const displayTitle = tocTitle || 'Table of Contents';

  // Filter headings based on user preferences
  const filteredHeadings = headings.filter(h => {
    if (h.level === 2 && !showH2) return false;
    if (h.level === 3 && !showH3) return false;
    if (h.level === 4 && !showH4) return false;
    return true;
  });

  if (filteredHeadings.length === 0) {
    return (
      <div className={styles.tocContainer}>
        <h3>{displayTitle}</h3>
        <p className={styles.tocEmpty}>
          No headings found. Add H2, H3, or H4 headings to your content.
        </p>
      </div>
    );
  }

  // Build nested structure for headings
  const buildNestedTOC = (items: IHeading[]): JSX.Element[] => {
    const result: JSX.Element[] = [];
    let i = 0;

    while (i < items.length) {
      const current = items[i];
      const children: IHeading[] = [];

      // Find all children (higher level numbers = sub-headings)
      let j = i + 1;
      while (j < items.length && items[j].level > current.level) {
        if (items[j].level === current.level + 1) {
          // Direct child
          children.push(items[j]);
        }
        j++;
      }

      const isActive = activeHeadingId === current.id;

      result.push(
        <li key={current.id}>
          <a
            href={`#${current.id}`}
            className={isActive ? styles.active : ''}
            onClick={(e: React.MouseEvent<HTMLAnchorElement>) => {
              e.preventDefault();
              onHeadingClick(current.id);
            }}
          >
            {current.text}
          </a>
          {children.length > 0 && <ul>{buildNestedTOC(children)}</ul>}
        </li>
      );

      // Skip children we've already processed
      i = j;
    }

    return result;
  };

  return (
    <div className={styles.tocContainer}>
      <h3>{displayTitle}</h3>
      <ul>
        {buildNestedTOC(filteredHeadings)}
      </ul>
    </div>
  );
};