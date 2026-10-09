import { IPropertyPaneGroup } from '@microsoft/sp-property-pane';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  FleetWebPartBase,
  IFleetWebPartProps,
} from '../../common/webpart/FleetWebPartBase';
import Reports, { IReportsProps } from './components/Reports';

export type IReportsWebPartProps = IFleetWebPartProps;

export default class ReportsWebPart extends FleetWebPartBase<IReportsWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IReportsProps> = React.createElement(Reports, {
      services: this.services,
      user: this.currentUser,
      bootstrapError: this.bootstrapError,
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

  protected settingsGroups(): IPropertyPaneGroup[] {
    // Reports read what the lists already hold. Nothing to configure.
    return [];
  }
}
