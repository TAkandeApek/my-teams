import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'MyTeamsHomeWebPartStrings';
import MyTeamsHome from './components/MyTeamsHome';
import { IMyTeamsHomeProps } from './components/IMyTeamsHomeProps';
import { MSGraphClient } from '@microsoft/sp-http';
import { TeamsService, ITeamsService } from '../../shared/services';

export interface IMyTeamsHomeWebPartProps {
  openInClientApp: boolean;
}

export default class MyTeamsHomeWebPart extends BaseClientSideWebPart <IMyTeamsHomeWebPartProps> {

  private _graphClient: MSGraphClient;
  private _teamsService: ITeamsService;

  public async onInit(): Promise<void> {
    if (DEBUG && Environment.type === EnvironmentType.Local) {
      console.log("Mock data service not implemented yet");
    } else {

      this._graphClient = await this.context.msGraphClientFactory.getClient();
      this._teamsService = new TeamsService(this._graphClient);
    }
    return super.onInit();
  }
  
   
  public render(): void {
    const element: React.ReactElement<IMyTeamsHomeProps> = React.createElement(
      MyTeamsHome,
      {
        teamsService: this._teamsService,
        openInClientApp: this.properties.openInClientApp
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
                PropertyPaneToggle('openInClientApp', {
                  label: strings.OpenInClientAppFieldLabel,
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
