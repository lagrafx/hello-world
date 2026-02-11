import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import DocumentCount, { IDocumentCountProps } from './components/DocumentCount';

export interface IDocumentCountWebPartProps {
  libraryTitle: string;
  refreshIntervalSeconds: number;
}

export default class DocumentCountWebPart extends BaseClientSideWebPart<IDocumentCountWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IDocumentCountProps> = React.createElement(DocumentCount, {
      libraryTitle: this.properties.libraryTitle,
      refreshIntervalSeconds: this.properties.refreshIntervalSeconds,
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl
    });

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
            description: 'Document count settings'
          },
          groups: [
            {
              groupName: 'Library',
              groupFields: [
                PropertyPaneTextField('libraryTitle', {
                  label: 'Library title',
                  placeholder: 'Documents'
                }),
                PropertyPaneSlider('refreshIntervalSeconds', {
                  label: 'Refresh interval (seconds)',
                  min: 10,
                  max: 300,
                  value: this.properties.refreshIntervalSeconds || 30,
                  showValue: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
