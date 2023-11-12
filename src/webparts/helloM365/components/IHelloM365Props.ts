import { Theme } from '@fluentui/react-components';
import { IMicrosoftTeams } from '@microsoft/sp-webpart-base';

export interface IHelloM365Props {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  teamsSdk?: IMicrosoftTeams;
  isGeoLocationAvailable: boolean;
  themeToApply: Theme;
  appContext: string;
  clientType: string;
}
