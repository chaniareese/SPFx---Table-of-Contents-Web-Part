export interface IHeading {
  id: string;
  text: string;
  level: number; // 2, 3, or 4 (H2, H3, H4)
  element?: HTMLElement; // Reference to the actual DOM element
}

export interface ITableOfContentsState {
  // List of headings extracted from content
  headings: IHeading[];
  
  // Currently active heading (for highlighting in TOC)
  activeHeadingId: string | null;
  
  // Content being edited
  editorContent: string;
}