import { IPropertyPaneGroup, PropertyPaneSlider } from '@microsoft/sp-property-pane';
import * as React from 'react';
import {
  FleetWebPartBase,
  IFleetWebPartProps,
} from '../../common/webpart/FleetWebPartBase';
import { FleetWebPartKey } from '../../common/webpart/access';
import DriverDashboard, {
  IDriverDashboardProps,
} from './components/DriverDashboard';

export interface IDriverDashboardWebPartProps extends IFleetWebPartProps {
  pollSeconds: number;
}

export default class DriverDashboardWebPart extends FleetWebPartBase<IDriverDashboardWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IDriverDashboardProps> = React.createElement(
      DriverDashboard,
      {
        services: this.services,
        user: this.currentUser,
        bootstrapError: this.bootstrapError,
        pollMs: (this.properties.pollSeconds ?? 20) * 1000,
      },
    );

    this.renderElement(element);
  }

  protected accessKey(): FleetWebPartKey {
    return 'driverDashboard';
  }

  protected settingsGroups(): IPropertyPaneGroup[] {
    return [
      {
        groupName: 'Offer polling',
        groupFields: [
          PropertyPaneSlider('pollSeconds', {
            label: 'Check for new offers every (seconds)',
            // The non-functional requirement: an open offer visible within 20
            // seconds of broadcast, at no more than one request per 20 seconds
            // per client. Going below 20 breaks the second half of that.
            min: 20,
            max: 120,
            step: 5,
          }),
        ],
      },
    ];
  }
}
