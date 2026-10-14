import { IPropertyPaneGroup, PropertyPaneSlider } from '@microsoft/sp-property-pane';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { APP_CONFIG } from '../../common/config/AppConfig';
import {
  FleetWebPartBase,
  IFleetWebPartProps,
} from '../../common/webpart/FleetWebPartBase';
import DispatchDashboard, {
  IDispatchDashboardProps,
} from './components/DispatchDashboard';

export interface IDispatchDashboardWebPartProps extends IFleetWebPartProps {
  refreshSeconds: number;
}

export default class DispatchDashboardWebPart extends FleetWebPartBase<IDispatchDashboardWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IDispatchDashboardProps> = React.createElement(
      DispatchDashboard,
      {
        services: this.services,
        user: this.currentUser,
        bootstrapError: this.bootstrapError,
        refreshMs: (this.properties.refreshSeconds ?? 60) * 1000,
        reportsUrl: `${this.context.pageContext.web.absoluteUrl}${APP_CONFIG.pages.reports}`,
      },
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

  protected settingsGroups(): IPropertyPaneGroup[] {
    return [
      {
        groupName: 'Dashboard',
        groupFields: [
          PropertyPaneSlider('refreshSeconds', {
            label: 'Refresh interval (seconds)',
            min: 30,
            max: 300,
            step: 15,
          }),
        ],
      },
    ];
  }
}
