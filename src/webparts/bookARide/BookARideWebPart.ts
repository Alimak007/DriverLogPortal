import {
  IPropertyPaneGroup,
  PropertyPaneSlider,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { APP_CONFIG } from '../../common/config/AppConfig';
import {
  FleetWebPartBase,
  IFleetWebPartProps,
} from '../../common/webpart/FleetWebPartBase';
import BookARide, { IBookARideProps } from './components/BookARide';

export interface IBookARideWebPartProps extends IFleetWebPartProps {
  defaultPropertyId: number;
  processingPollMs: number;
}

export default class BookARideWebPart extends FleetWebPartBase<IBookARideWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IBookARideProps> = React.createElement(BookARide, {
      services: this.services,
      user: this.currentUser,
      bootstrapError: this.bootstrapError,
      // The pane field is a text box, so it arrives as a string, not a number.
      defaultPropertyId: parseInt(String(this.properties.defaultPropertyId ?? ''), 10) || 0,
      processingPollMs: (this.properties.processingPollMs ?? 10) * 1000,
      myBookingsUrl: `${this.context.pageContext.web.absoluteUrl}${APP_CONFIG.pages.myBookings}`,
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

  protected settingsGroups(): IPropertyPaneGroup[] {
    return [
      {
        groupName: 'Booking',
        groupFields: [
          PropertyPaneTextField('defaultPropertyId', {
            label: 'Default property ID',
            description:
              'The PropertyMaster item ID this page books against. Leave at 0 to use ' +
              "the signed-in user's own property.",
          }),
          PropertyPaneSlider('processingPollMs', {
            label: 'Processing screen refresh (seconds)',
            min: 5,
            max: 60,
            step: 5,
          }),
        ],
      },
    ];
  }
}
