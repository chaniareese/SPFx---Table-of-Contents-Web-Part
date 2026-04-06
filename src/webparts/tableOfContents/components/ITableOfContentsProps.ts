import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ITableOfContentsProps {
  // Content stored in web part properties
  content: string;
  
  // Title of the runbook
  title: string;
  
  // Toggle options for which headings to show in TOC
  showH2: boolean;
  showH3: boolean;
  showH4: boolean;
  
  // Whether to hide in mobile view
  hideInMobile: boolean;
  
  // Callback to update properties when content changes
  updateProperty: (value: string) => void;
  
  // Display mode (Edit or Read)
  displayMode: number;

  // SharePoint context — needed for ReviewDateNode, LastUpdatedNode, VersionNode
  siteUrl: string;
  pageId: number | undefined;
  context: WebPartContext;
}