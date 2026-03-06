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
  private updateTimeout: NodeJS.Timeout | null = null;

  constructor(props: ITableOfContentsProps) {
    super(props);
    this.state = {
      headings: [],
      activeHeadingId: null,
      editorContent: props.content || ''
    };
  }

  public componentDidMount(): void {
    // Extract headings after component mounts
    this.extractHeadings();
    
    // Set up scroll listener for active heading detection
    window.addEventListener('scroll', this.handleScroll);
  }

  public componentWillUnmount(): void {
    window.removeEventListener('scroll', this.handleScroll);
    if (this.updateTimeout) {
      clearTimeout(this.updateTimeout);
    }
  }

  public componentDidUpdate(prevProps: ITableOfContentsProps): void {
    // Re-extract headings if content changes
    if (prevProps.content !== this.props.content) {
      this.extractHeadings();
      this.setState({ editorContent: this.props.content });
    }
  }

  // Extract H2, H3, H4 headings from content
  private extractHeadings = (): void => {
    if (!this.contentRef.current) return;

    const headingElements = this.contentRef.current.querySelectorAll('h2, h3, h4');
    const headings: IHeading[] = [];

    headingElements.forEach((element: Element, index: number) => {
      const htmlElement = element as HTMLElement;
      const level = parseInt(element.tagName.substring(1)); // H2 -> 2, H3 -> 3
      const text = htmlElement.innerText || htmlElement.textContent || '';
      
      // Generate unique ID if not present
      let id = htmlElement.id;
      if (!id) {
        id = `heading-${level}-${index}`;
        htmlElement.id = id;
      }

      headings.push({
        id,
        text: text.trim(),
        level,
        element: htmlElement
      });
    });

    this.setState({ headings });
  };

  // Detect which heading is currently in view
  private handleScroll = (): void => {
    const { headings } = this.state;
    if (headings.length === 0) return;

    // Find the heading closest to the top of the viewport
    let activeId: string | null = null;
    const scrollPosition = window.scrollY + 150; // Offset for better UX

    for (let i = headings.length - 1; i >= 0; i--) {
      const heading = headings[i];
      if (heading.element) {
        const rect = heading.element.getBoundingClientRect();
        const elementTop = rect.top + window.scrollY;
        
        if (scrollPosition >= elementTop) {
          activeId = heading.id;
          break;
        }
      }
    }

    if (activeId !== this.state.activeHeadingId) {
      this.setState({ activeHeadingId: activeId });
    }
  };

  // Handle clicking on TOC links
  private handleTOCClick = (id: string): void => {
    const element = document.getElementById(id);
    if (element) {
      element.scrollIntoView({ 
        behavior: 'smooth', 
        block: 'start' 
      });
      
      // Update active heading
      this.setState({ activeHeadingId: id });
    }
  };

  // Handle content changes in the editor
  private handleEditorChange = (content: string): void => {
    this.setState({ editorContent: content });

    // Debounce the update to avoid too many calls
    if (this.updateTimeout) {
      clearTimeout(this.updateTimeout);
    }

    this.updateTimeout = setTimeout(() => {
      this.props.updateProperty(content);
      // Re-extract headings after content updates
      setTimeout(() => this.extractHeadings(), 100);
    }, 500);
  };

  // Quill editor configuration
  private modules = {
    toolbar: [
      [{ 'header': [2, 3, 4, false] }],
      ['bold', 'italic', 'underline', 'strike'],
      [{ 'list': 'ordered'}, { 'list': 'bullet' }],
      [{ 'indent': '-1'}, { 'indent': '+1' }],
      ['link', 'image', 'code-block'],
      [{ 'color': [] }, { 'background': [] }],
      ['clean']
    ]
  };

  private formats = [
    'header',
    'bold', 'italic', 'underline', 'strike',
    'list', 'bullet', 'indent',
    'link', 'image', 'code-block',
    'color', 'background'
  ];

  public render(): React.ReactElement<ITableOfContentsProps> {
    const { title, showH2, showH3, showH4, hideInMobile, displayMode } = this.props;
    const { headings, activeHeadingId, editorContent } = this.state;

    const isEditMode = displayMode === DisplayMode.Edit;
    const containerClass = hideInMobile ? `${styles.tableOfContents} ${styles.hideInMobile}` : styles.tableOfContents;

    // Empty state when no content
    if (!editorContent && !isEditMode) {
      return (
        <div className={styles.emptyState}>
          <h3>📋 No Content Yet</h3>
          <p>Edit this page and add content to your runbook to get started.</p>
        </div>
      );
    }

    return (
      <div className={containerClass}>
        {/* Table of Contents - Sticky on left */}
        <TOCNavigator
          headings={headings}
          activeHeadingId={activeHeadingId}
          showH2={showH2}
          showH3={showH3}
          showH4={showH4}
          onHeadingClick={this.handleTOCClick}
        />

        {/* Content Area - Scrollable on right */}
        <div className={styles.contentContainer}>
          {title && <h1 className={styles.title}>{title}</h1>}

          {isEditMode ? (
            // Edit mode: Show rich text editor
            <div className={styles.editorContainer}>
              <ReactQuill
                value={editorContent}
                onChange={this.handleEditorChange}
                modules={this.modules}
                formats={this.formats}
                theme="snow"
                placeholder="Start writing your DR/BC runbook content here... Use Heading 2, 3, or 4 for sections that will appear in the Table of Contents."
              />
            </div>
          ) : (
            // View mode: Display content
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