import * as React from 'react';
import styles from './TableOfContents.module.scss';
import { ITableOfContentsProps } from './ITableOfContentsProps';
import { ITableOfContentsState, IHeading } from './ITableOfContentsState';
import { TOCNavigator } from './TOCNavigator';
import ReactQuill from 'react-quill';
import 'quill/dist/quill.snow.css';
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
    const editContainer = this.editorContainerRef.current?.querySelector('.ql-editor') as HTMLElement | null;
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
    const contentEl = this.contentContainerRef.current;
    if (headings.length === 0 || !contentEl) return;

    let activeId: string | null = null;
    const containerTop = contentEl.getBoundingClientRect().top;

    for (let i = headings.length - 1; i >= 0; i--) {
      const heading = headings[i];
      if (heading.element) {
        const rect = heading.element.getBoundingClientRect();
        if (rect.top - containerTop <= 80) {
          activeId = heading.id;
          break;
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

  private modules = {
    toolbar: [
      [{ 'header': [2, 3, 4, false] }],
      ['bold', 'italic', 'underline', 'strike'],
      [{ 'list': 'ordered' }, { 'list': 'bullet' }],
      [{ 'indent': '-1' }, { 'indent': '+1' }],
      ['link', 'image', 'code-block'],
      [{ 'color': [] }, { 'background': [] }],
      ['clean']
    ]
  };

  private formats = [
    'header', 'bold', 'italic', 'underline', 'strike',
    'list', 'bullet', 'indent',
    'link', 'image', 'code-block', 'color', 'background'
  ];

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
          onHeadingClick={this.handleTOCClick}
        />
        <div className={styles.contentContainer} ref={this.contentContainerRef}>
          {title && <h1 className={styles.title}>{title}</h1>}
          {isEditMode ? (
            <div className={styles.editorContainer} ref={this.editorContainerRef}>
              <ReactQuill
                value={editorContent}
                onChange={this.handleEditorChange}
                modules={this.modules}
                formats={this.formats}
                theme="snow"
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