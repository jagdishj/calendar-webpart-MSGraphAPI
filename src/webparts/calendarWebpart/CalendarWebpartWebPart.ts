import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'CalendarWebpartWebPartStrings';
import CalendarWebpart from './components/CalendarWebpart';
import { ICalendarWebpartProps } from './components/ICalendarWebpartProps';
import { MSGraphClient } from '@microsoft/sp-http';
export interface ICalendarWebpartWebPartProps {
  description: string;
}

export default class CalendarWebpartWebPart extends BaseClientSideWebPart<ICalendarWebpartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICalendarWebpartProps > = React.createElement(
      CalendarWebpart,
      {
        description: this.properties.description,
        context:this.context
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
