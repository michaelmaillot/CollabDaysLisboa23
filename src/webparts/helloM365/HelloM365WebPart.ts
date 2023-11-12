import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'HelloM365WebPartStrings';
import HelloM365 from './components/HelloM365';
import { IHelloM365Props } from './components/IHelloM365Props';
import {
  teamsLightTheme,
  webLightTheme,
  webDarkTheme,
  teamsDarkTheme,
  Theme,
} from '@fluentui/react-components';
import { createV9Theme } from "@fluentui/react-migration-v8-v9";

export interface IHelloM365WebPartProps {
  description: string;
}

export default class HelloM365WebPart extends BaseClientSideWebPart<IHelloM365WebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _environment: string = 'SharePoint Online';
  private _clientType: string = 'web';
  private _theme: Theme = webLightTheme;

  public render(): void {
    const element: React.ReactElement<IHelloM365Props> = React.createElement(
      HelloM365,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        teamsSdk: this.context.sdks.microsoftTeams,
        isGeoLocationAvailable: !!this.context.sdks.microsoftTeams && this.context.sdks.microsoftTeams.teamsJs.geoLocation.isSupported(),
        themeToApply: this._theme,
        appContext: this._environment,
        clientType: this._clientType
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;

if (!!this.context.sdks.microsoftTeams) {
      switch (this._environment) {
        case 'Teams':
          if (this._isDarkTheme) {
            this._theme = teamsDarkTheme;
          }
          else {
            this._theme = teamsLightTheme;
          }
          break;

        default:
          if (this._isDarkTheme) {
            this._theme = webDarkTheme;
          }
          else {
            this._theme = webLightTheme;
          }
      }
    }
    else {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      this._theme = createV9Theme(currentTheme as any, webLightTheme);
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
