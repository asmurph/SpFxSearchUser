import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SpFxSearchUserWebPartStrings';
import SpFxSearchUser from './components/SpFxSearchUser';
import { ISpFxSearchUserProps } from './components/ISpFxSearchUserProps';

export interface ISpFxSearchUserWebPartProps {
  description: string;
}

export default class SpFxSearchUserWebPart extends BaseClientSideWebPart<ISpFxSearchUserWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpFxSearchUserProps> = React.createElement(
      SpFxSearchUser,
      {
        context: this.context,
        description: this.properties.description
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
