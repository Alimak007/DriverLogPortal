import { IPropertyPaneGroup, PropertyPaneSlider } from '@microsoft/sp-property-pane';
import * as React from 'react';
import { APP_CONFIG } from '../../common/config/AppConfig';
import {
  FleetWebPartBase,
  IFleetWebPartProps,
} from '../../common/webpart/FleetWebPartBase';
import { FleetWebPartKey } from '../../common/webpart/access';
import MyBookings, { IMyBookingsProps } from './components/MyBookings';

export interface IMyBookingsWebPartProps extends IFleetWebPartProps {
  historyCount: number;
  refreshSeconds: number;
}

export default class MyBookingsWebPart extends FleetWebPartBase<IMyBookingsWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IMyBookingsProps> = React.createElement(MyBookings, {
      services: this.services,
      user: this.currentUser,
      bootstrapError: this.bootstrapError,
      historyCount: this.properties.historyCount ?? 50,
      refreshMs: (this.properties.refreshSeconds ?? 30) * 1000,
      bookARideUrl: `${this.context.pageContext.web.absoluteUrl}${APP_CONFIG.pages.bookARide}`,
    });

    this.renderElement(element);
  }

  protected accessKey(): FleetWebPartKey {
    return 'myBookings';
  }

  protected settingsGroups(): IPropertyPaneGroup[] {
    return [
      {
        groupName: 'History',
        groupFields: [
          PropertyPaneSlider('historyCount', {
            label: 'Requests to show',
            min: 10,
            max: 100,
            step: 10,
          }),
          PropertyPaneSlider('refreshSeconds', {
            label: 'Refresh interval (seconds)',
            min: 15,
            max: 300,
            step: 15,
          }),
        ],
      },
    ];
  }
}
