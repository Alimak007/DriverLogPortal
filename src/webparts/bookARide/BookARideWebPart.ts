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
  googleMapsApiKey: string;
  googleMapsMapId: string;
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
      googleMapsApiKey:
        String(this.properties.googleMapsApiKey ?? '').trim() || APP_CONFIG.googleMaps.apiKey,
      googleMapsMapId:
        String(this.properties.googleMapsMapId ?? '').trim() || APP_CONFIG.googleMaps.mapId,
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
          PropertyPaneTextField('googleMapsApiKey', {
            label: 'Google Maps API key',
            description:
              'Overrides the shared key in APP_CONFIG for this site. Either way ' +
              'the key ships in the bundle and IS extractable, so it must be ' +
              'restricted by HTTP referrer AND by API (Maps JavaScript, Places ' +
              'New, Geocoding) in the Google Cloud console, with a budget alert ' +
              'set. This key is billed, so an unrestricted one is a bill, not a ' +
              'rate limit.',
          }),
          PropertyPaneTextField('googleMapsMapId', {
            label: 'Google Maps Map ID',
            description:
              'Optional. Enables Advanced Markers and cloud-based map styling. ' +
              'Without one the map still works but uses the deprecated classic ' +
              'marker. Create it under Google Maps Platform > Map management.',
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
