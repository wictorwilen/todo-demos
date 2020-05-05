import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';

import * as strings from 'ToDoWebPartStrings';
import ToDo from './components/ToDo';
import { IToDoProps } from './components/IToDoProps';

export interface IToDoWebPartProps {
  itemCount: number;
}

export default class ToDoWebPart extends BaseClientSideWebPart<IToDoWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IToDoProps> = React.createElement(
      ToDo,
      {
        itemCount: this.properties.itemCount,
        context: this.context
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
                PropertyPaneSlider('itemCount', {
                  label: strings.ItemCountFieldLabel,
                  min: 1,
                  max: 100
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
