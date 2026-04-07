import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
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
      editorContent: props.content || '',
      viewContent: props.content || ''
    };
  }

  public componentDidMount(): void {
    this.extractHeadings();
    this.buildViewContent();
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
      this.buildViewContent();
    }
  }


  // Build view mode content — replaces dynamic node spans with live data from RunbookReviews
  private buildViewContent(): void {
    const { siteUrl, pageId, content } = this.props;
    if (!siteUrl || !pageId) {
      this.setState({ viewContent: content });
      return;
    }

    const url = `${siteUrl}/_api/web/lists/getbytitle('RunbookReviews')/items` +
      `?$select=ReviewedDate,NextReviewDate` +
      `&$filter=RunbookPageId eq ${pageId}` +
      `&$orderby=ReviewedDate desc` +
      `&$top=500`;

    this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((r: SPHttpClientResponse) => r.json())
      .then((data: { value: { ReviewedDate: string; NextReviewDate: string }[] }) => {
        const items = data.value || [];
        const count = items.length;
        const latest = count > 0 ? items[0] : null;

        const formatDate = (iso: string): string => {
          if (!iso) return '';
          const d = new Date(iso);
          if (isNaN(d.getTime())) return '';
          return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
        };

        // Replace dynamic node spans in HTML with actual values
        let html = content;

        if (latest) {
          const lastUpdated = formatDate(latest.ReviewedDate);
          const nextReview = formatDate(latest.NextReviewDate);
          const version = `v${count + 1}.0`;

          // Replace last-updated spans
          html = html.replace(
            /<span[^>]*data-type="last-updated"[^>]*data-fallback="([^"]*)"[^>]*><\/span>/gi,
            lastUpdated || '$1'
          );
          // Also handle reversed attribute order
          html = html.replace(
            /<span[^>]*data-fallback="([^"]*)"[^>]*data-type="last-updated"[^>]*><\/span>/gi,
            lastUpdated || '$1'
          );

          // Replace review-date spans
          html = html.replace(
            /<span[^>]*data-type="review-date"[^>]*data-fallback="([^"]*)"[^>]*><\/span>/gi,
            nextReview || '$1'
          );
          html = html.replace(
            /<span[^>]*data-fallback="([^"]*)"[^>]*data-type="review-date"[^>]*><\/span>/gi,
            nextReview || '$1'
          );

          // Replace version spans
          html = html.replace(
            /<span[^>]*data-type="version"[^>]*data-fallback="([^"]*)"[^>]*><\/span>/gi,
            version
          );
          html = html.replace(
            /<span[^>]*data-fallback="([^"]*)"[^>]*data-type="version"[^>]*><\/span>/gi,
            version
          );
        } else {
          // No review records — replace spans with their fallback values
          html = html.replace(
            /<span[^>]*data-type="(?:last-updated|review-date|version)"[^>]*data-fallback="([^"]*)"[^>]*><\/span>/gi,
            '$1'
          );
          html = html.replace(
            /<span[^>]*data-fallback="([^"]*)"[^>]*data-type="(?:last-updated|review-date|version)"[^>]*><\/span>/gi,
            '$1'
          );
        }

        this.setState({ viewContent: html });
      })
      .catch(() => {
        this.setState({ viewContent: content });
      });
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

      // Generate stable ID based on heading content + level + index
      // Use data attribute to survive TipTap DOM re-renders
      const stableId = `toc-${level}-${index}-${text.trim().toLowerCase().replace(/[^a-z0-9]+/g, '-')}`;
      const existingId = htmlElement.getAttribute('data-toc-id');
      const id = existingId || stableId;
      
      if (!existingId) {
        htmlElement.setAttribute('data-toc-id', id);
        // Also set DOM id as fallback for direct getElementById calls
        if (!htmlElement.id) {
          htmlElement.id = id;
        }
      }

      // Don't store element reference - it becomes stale after TipTap re-renders
      // We'll look up elements fresh on each scroll event
      headings.push({ id, text: text.trim(), level, element: undefined });
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
    let activeId: string | null = null;

    // Strategy: find the LAST heading that has scrolled past the top of the viewport.
    // No lower bound on distance — long sections can push headings hundreds of pixels
    // above the viewport and they should still remain active.
    let lastScrolledPastIndex = -1;

    for (let i = 0; i < headings.length; i++) {
      const heading = headings[i];
      const element = contentEl.querySelector(`[data-toc-id="${heading.id}"]`) as HTMLElement | null;
      if (element) {
        const rect = element.getBoundingClientRect();
        const distanceFromTop = rect.top - containerTop;
        // Heading has scrolled past the top threshold (80px offset)
        if (distanceFromTop <= 80) {
          lastScrolledPastIndex = i;
        }
      }
    }

    if (lastScrolledPastIndex !== -1) {
      const activeHeading = headings[lastScrolledPastIndex];
      if (isVisible(activeHeading.level)) {
        activeId = activeHeading.id;
      } else {
        // Walk backwards to find nearest visible heading above
        for (let i = lastScrolledPastIndex - 1; i >= 0; i--) {
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
    // Look for element by data-toc-id first (survives re-renders), fall back to id
    let target = contentEl?.querySelector(`[data-toc-id="${id}"]`) as HTMLElement | null;
    if (!target) {
      target = document.getElementById(id);
    }
    if (contentEl && target) {
      const scrollTop = contentEl.scrollTop;
      const containerTop = contentEl.getBoundingClientRect().top;
      const targetTop = target.getBoundingClientRect().top;
      const absoluteTargetTop = scrollTop + targetTop - containerTop - 100;
      contentEl.scrollTo({ top: absoluteTargetTop, behavior: 'smooth' });
      this.setState({ activeHeadingId: id });
      // Trigger scroll detection after scroll completes
      setTimeout(() => this.handleScroll(), 800);
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
                siteUrl={this.props.siteUrl}
                pageId={this.props.pageId}
              />
            </div>
          ) : (
            <div
              className={styles.contentDisplay}
              ref={this.contentRef}
              dangerouslySetInnerHTML={{ __html: this.state.viewContent }}
            />
          )}
        </div>
      </div>
    );
  }
}