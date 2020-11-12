import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'TestWebPartStrings';
import Test from './components/Test';
import { ITestProps } from './components/ITestProps';

export interface ITestWebPartProps {
  description: string;
}

export default class TestWebPart extends BaseClientSideWebPart<ITestWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITestProps> = React.createElement(
      Test,
      {
        description: this.properties.description,
        spHttpClient: this.context.spHttpClient,  
        siteUrl: this.context.pageContext.web.absoluteUrl ,
        listName: "LearningSession"
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
