import * as React from 'react';
import styles from './TableOfContents.module.scss';
import { ITableOfContentsProps } from './ITableOfContentsProps';
import { ITableOfContentsState, IHeading } from './ITableOfContentsState';
import { TOCNavigator } from './TOCNavigator';
import { TipTapEditor } from './TipTapEditor';
import { DisplayMode } from '@microsoft/sp-core-library';

export default class TableOfContents extends React.Component<ITableOfContentsProps, ITableOfContentsState> {
  private contentRef = React.createRef<HTMLDivElement>();
  private editorContainerRef = React.createRef<HTMLDivElement>();
  private contentContainerRef = React.createRef<HTMLDivElement>();

  private updateTimeout: ReturnType<typeof setTimeout> | null = null;
  private scanInterval: ReturnType<typeof setInterval> | null = null;

  constructor(props: ITableOfContentsProps) {
    super(props);
    this.state = {
      headings: [],
      activeHeadingId: null,
      editorContent: props.content || ''
    };
  }

  public componentDidMount(): void {
    this.extractHeadings();
    this.scanInterval = setInterval(() => this.extractHeadings(), 2000);
    const contentEl = this.contentContainerRef.current;
    if (contentEl) {
      contentEl.addEventListener('scroll', this.handleScroll);
    }
  }

  public componentWillUnmount(): void {
    const contentEl = this.contentContainerRef.current;
    if (contentEl) {
      contentEl.removeEventListener('scroll', this.handleScroll);
    }
    if (this.updateTimeout) clearTimeout(this.updateTimeout);
    if (this.scanInterval) clearInterval(this.scanInterval);
  }

  public componentDidUpdate(prevProps: ITableOfContentsProps): void {
    if (prevProps.content !== this.props.content) {
      this.extractHeadings();
      this.setState({ editorContent: this.props.content });
    }
  }

  private extractHeadings = (): void => {
    const isEdit = this.props.displayMode === DisplayMode.Edit;
    const readContainer = this.contentRef.current;
    const editContainer = this.editorContainerRef.current?.querySelector('.ProseMirror') as HTMLElement | null;
    const headingContainer = isEdit ? editContainer : readContainer;

    if (!headingContainer) {
      this.setState({ headings: [] });
      return;
    }

    const headingElements = headingContainer.querySelectorAll('h2, h3, h4');
    const headings: IHeading[] = [];

    headingElements.forEach((element: Element, index: number) => {
      const htmlElement = element as HTMLElement;
      const level = parseInt(element.tagName.substring(1), 10);
      const text = htmlElement.innerText || htmlElement.textContent || '';
      if (!text.trim()) return;

      let id = htmlElement.id;
      if (!id) {
        id = `heading-${level}-${index}`;
        htmlElement.id = id;
      }

      headings.push({ id, text: text.trim(), level, element: htmlElement });
    });

    const nextSig = headings.map(h => `${h.id}|${h.level}|${h.text}`).join('||');
    const prevSig = this.state.headings.map(h => `${h.id}|${h.level}|${h.text}`).join('||');
    if (nextSig !== prevSig) {
      this.setState({ headings });
    }
  };

  private handleScroll = (): void => {
    const { headings } = this.state;
    const { showH2, showH3, showH4 } = this.props;
    const contentEl = this.contentContainerRef.current;
    if (headings.length === 0 || !contentEl) return;

    const isVisible = (level: number): boolean => {
      if (level === 2) return showH2;
      if (level === 3) return showH3;
      if (level === 4) return showH4;
      return true;
    };

    const containerTop = contentEl.getBoundingClientRect().top;

    // Find the heading the user has scrolled past (regardless of visibility)
    let scrolledPastIndex = -1;
    for (let i = headings.length - 1; i >= 0; i--) {
      const heading = headings[i];
      if (heading.element) {
        const rect = heading.element.getBoundingClientRect();
        if (rect.top - containerTop <= 80) {
          scrolledPastIndex = i;
          break;
        }
      }
    }

    let activeId: string | null = null;

    if (scrolledPastIndex !== -1) {
      // If the scrolled-past heading is visible, use it directly
      if (isVisible(headings[scrolledPastIndex].level)) {
        activeId = headings[scrolledPastIndex].id;
      } else {
        // Otherwise walk backwards to find the nearest visible ancestor/sibling above
        for (let i = scrolledPastIndex - 1; i >= 0; i--) {
          if (isVisible(headings[i].level)) {
            activeId = headings[i].id;
            break;
          }
        }
      }
    }

    if (activeId !== this.state.activeHeadingId) {
      this.setState({ activeHeadingId: activeId });
    }
  };

  private handleTOCClick = (id: string): void => {
    const contentEl = this.contentContainerRef.current;
    const target = document.getElementById(id);
    if (contentEl && target) {
      const scrollTop = contentEl.scrollTop;
      const containerTop = contentEl.getBoundingClientRect().top;
      const targetTop = target.getBoundingClientRect().top;
      const absoluteTargetTop = scrollTop + targetTop - containerTop - 10;
      contentEl.scrollTo({ top: absoluteTargetTop, behavior: 'smooth' });
      this.setState({ activeHeadingId: id });
    }
  };

  private handleEditorChange = (content: string): void => {
    this.setState({ editorContent: content });
    if (this.updateTimeout) clearTimeout(this.updateTimeout);
    this.updateTimeout = setTimeout(() => {
      this.props.updateProperty(content);
      setTimeout(() => this.extractHeadings(), 100);
    }, 500);
  };

  public render(): React.ReactElement<ITableOfContentsProps> {
    const { title, showH2, showH3, showH4, hideInMobile, displayMode } = this.props;
    const { headings, activeHeadingId, editorContent } = this.state;
    const isEditMode = displayMode === DisplayMode.Edit;

    const containerClass = hideInMobile
      ? `${styles.tableOfContents} ${styles.hideInMobile}`
      : styles.tableOfContents;

    if (!editorContent && !isEditMode) {
      return (
        <div className={styles.emptyState}>
          <h3>📋 No Content Yet</h3>
          <p>Edit this page and add content to get started.</p>
        </div>
      );
    }

    return (
      <div className={containerClass}>
        <TOCNavigator
          headings={headings}
          activeHeadingId={activeHeadingId}
          showH2={showH2}
          showH3={showH3}
          showH4={showH4}
          tocTitle={title}
          onHeadingClick={this.handleTOCClick}
        />
        <div className={styles.contentContainer} ref={this.contentContainerRef}>
          {isEditMode ? (
            <div className={styles.editorContainer} ref={this.editorContainerRef}>
              <TipTapEditor
                content={editorContent}
                onChange={this.handleEditorChange}
                placeholder="Start writing your content here… Use Heading 2, 3, or 4 for sections that will appear in the Table of Contents."
              />
            </div>
          ) : (
            <div
              className={styles.contentDisplay}
              ref={this.contentRef}
              dangerouslySetInnerHTML={{ __html: editorContent }}
            />
          )}
        </div>
      </div>
    );
  }
}