import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'TableOfContentsWebPartStrings';
import TableOfContents from './components/TableOfContents';
import { ITableOfContentsProps } from './components/ITableOfContentsProps';

export interface ITableOfContentsWebPartProps {
  title: string;
  content: string;
  showH2: boolean;
  showH3: boolean;
  showH4: boolean;
  hideInMobile: boolean;
}

export default class TableOfContentsWebPart extends BaseClientSideWebPart<ITableOfContentsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITableOfContentsProps> = React.createElement(
      TableOfContents,
      {
        title: this.properties.title,
        content: this.properties.content,
        showH2: this.properties.showH2,
        showH3: this.properties.showH3,
        showH4: this.properties.showH4,
        hideInMobile: this.properties.hideInMobile,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.content = value;
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure your Table of Contents settings'
          },
          groups: [
            {
              groupName: 'General Settings',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'TOC Title',
                  placeholder: 'e.g., Table of Contents'
                })
              ]
            },
            {
              groupName: 'Table of Contents',
              groupFields: [
                PropertyPaneToggle('showH2', {
                  label: 'Show Heading 2 (H2)',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('showH3', {
                  label: 'Show Heading 3 (H3)',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('showH4', {
                  label: 'Show Heading 4 (H4)',
                  onText: 'Visible',
                  offText: 'Hidden'
                })
              ]
            },
            {
              groupName: 'Display Options',
              groupFields: [
                PropertyPaneToggle('hideInMobile', {
                  label: 'Hide TOC on Mobile',
                  onText: 'Hidden',
                  offText: 'Visible'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  // Set default property values
  protected onInit(): Promise<void> {
    // Set defaults if properties are undefined
    if (this.properties.showH2 === undefined) {
      this.properties.showH2 = true;
    }
    if (this.properties.showH3 === undefined) {
      this.properties.showH3 = true;
    }
    if (this.properties.showH4 === undefined) {
      this.properties.showH4 = true;
    }
    if (this.properties.hideInMobile === undefined) {
      this.properties.hideInMobile = false;
    }
    if (!this.properties.content) {
      this.properties.content = '<h2>Getting Started</h2><p>Welcome to your DR/BC System Runbook! Start editing to add your content.</p><h3>Quick Start</h3><p>Use Heading 2 for main sections and Heading 3 for subsections.</p>';
    }
    if (!this.properties.title) {
      this.properties.title = 'Table of Contents';
    }

    return super.onInit();
  }
}
