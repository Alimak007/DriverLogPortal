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
  tomtomApiKey: string;
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
      // The pane wins where a site has its own key; otherwise the shared
      // default in APP_CONFIG applies, so the form works with no setup.
      tomtomApiKey:
        String(this.properties.tomtomApiKey ?? '').trim() || APP_CONFIG.tomtom.apiKey,
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
          PropertyPaneTextField('tomtomApiKey', {
            label: 'TomTom API key',
            description:
              'Optional. Overrides the shared key in APP_CONFIG for this site. ' +
              'Either way the key ships in the bundle and IS extractable, so it ' +
              'must be restricted by HTTP referrer in the TomTom portal and have ' +
              'a daily usage alert set.',
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
