import * as React from 'react';
import styles from './HelloM365.module.scss';
import { IHelloM365Props } from './IHelloM365Props';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  FluentProvider,
  Button,
  Text,
  TabList,
  Tab,
  Checkbox,
  SelectTabEvent,
  SelectTabData,
  TabValue
} from '@fluentui/react-components';
import { geoLocation } from '@microsoft/teams-js';

interface IHelloM365State {
  coordinates: geoLocation.Location | undefined;
  selectedTab: TabValue;
}

export default class HelloM365 extends React.Component<IHelloM365Props, IHelloM365State> {
  public constructor(props: IHelloM365Props) {
    super(props);
    this.state = {
      coordinates: undefined,
      selectedTab: "App"
    };
  }

  public async componentDidMount(): Promise<void> {
    if (this.props.hasTeamsContext && this.props.teamsSdk) {
      if (this.props.teamsSdk.teamsJs.app.isInitialized()) {
        try {

          if (this.props.teamsSdk.teamsJs.geoLocation.isSupported()) {
            if (!(await this.props.teamsSdk.teamsJs.geoLocation.hasPermission())) {
              await this.props.teamsSdk.teamsJs.geoLocation.requestPermission();
            }
          }

          if (this.props.teamsSdk.teamsJs.pages.appButton.isSupported()) {
            this.props.teamsSdk.teamsJs.pages.appButton.onHoverEnter(() => {
              console.log("Hover enter");
            })
          }
        } catch (error) {
          console.log(error);
        }
      }
    }
  }

  public render(): React.ReactElement<IHelloM365Props> {
    const {
      isDarkTheme,
      environmentMessage,
      userDisplayName,
      isGeoLocationAvailable,
      teamsSdk,
      themeToApply,
      hasTeamsContext,
      clientType
    } = this.props;

    return (
      <FluentProvider theme={themeToApply}>
        <div className={styles.helloM365}>
          <div className={styles.welcome}>
            <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
            <h2>Well done, {escape(userDisplayName)}!</h2>
            <div id="envMessage">{environmentMessage}</div>
            <div>Client type: {clientType}</div>
          </div>
          <div>
            <h3>Welcome to SharePoint Framework! Yeah!</h3>
            <Button appearance="primary">A simple button</Button>
            <TabList selectedValue={this.state.selectedTab} disabled={!hasTeamsContext} onTabSelect={this._tabClickHandler}>
              <Tab value="App">
                App
              </Tab>
              <Tab value="Pages">
                Pages
              </Tab>
              <Tab value="GeoLocation">
                GeoLocation
              </Tab>
              <Tab value="AppInstallDialog">
                AppInstallDialog
              </Tab>
              <Tab value="Mail">
                Mail
              </Tab>
            </TabList>
            <div style={{ paddingTop: '15px' }}>
              {this.state.selectedTab === "Pages" &&
                <div>
                  <Checkbox disabled label="pages" checked={teamsSdk?.teamsJs.pages.isSupported()} />
                  <Checkbox disabled label="appButton" checked={teamsSdk?.teamsJs.pages.appButton.isSupported()} />
                  <Checkbox disabled label="backStack" checked={teamsSdk?.teamsJs.pages.backStack.isSupported()} />
                </div>
              }
              {this.state.selectedTab === "GeoLocation" &&
                <div>
                  <Checkbox disabled label="geolocation.map" checked={teamsSdk?.teamsJs.geoLocation.map.isSupported()} />
                  <Checkbox disabled label="geolocation" checked={isGeoLocationAvailable} />
                  {isGeoLocationAvailable &&
                    <div style={{ columnGap: "15px", display: "flex" }}>
                      <Button
                        onClick={async () => {
                          try {
                            const location = await teamsSdk?.teamsJs.geoLocation.getCurrentLocation();
                            console.log(location);
                            this.setState({ coordinates: location });
                          } catch (e) {
                            console.log(e);
                            console.log(`GeoLocation error: ${e}`);
                          }
                        }}
                      >
                        Get current location
                      </Button>
                      {this.state.coordinates &&
                        <Text style={{ marginTop: '5px' }}>{`Accuracy: ${this.state.coordinates.accuracy} - Latitude: ${this.state.coordinates.latitude} - Longitude: ${this.state.coordinates.longitude}`}</Text>
                      }
                    </div>
                  }
                </div>
              }
              {this.state.selectedTab === "AppInstallDialog" &&
                <div>
                  <Checkbox disabled label="appInstallDialog" checked={teamsSdk?.teamsJs.appInstallDialog.isSupported()} />
                  <Button appearance="primary" disabled={!teamsSdk?.teamsJs.appInstallDialog.isSupported()}
                    onClick={async () => {
                      try {
                        await teamsSdk?.teamsJs.appInstallDialog.openAppInstallDialog({
                          appId: "018ce865-febc-41bb-99f8-651e5abd18e0",
                        });

                        if (teamsSdk?.teamsJs.pages.currentApp.isSupported()) {
                          // await teamsSdk.teamsJs.pages.currentApp.navigateTo({
                          //   pageId: "about"
                          // });

                          // await teamsSdk.teamsJs.pages.navigateToApp({
                          //   appId: "57e078b5-6c0e-44a1-a83f-45f75b030d4a",
                          //   pageId: "sections/MyAssist"
                          // });
                        }
                      } catch (e) {
                        console.log(e);
                        console.log(`App navigation error: ${e}`);
                      }
                    }}
                  >
                    Open App Info
                  </Button>
                </div>
              }
              {this.state.selectedTab === "App" && hasTeamsContext &&
                <div style={{ columnGap: "15px", display: "flex" }}>
                  <Button appearance="primary"
                    onClick={async () => {
                      await teamsSdk?.teamsJs.app.openLink("https://teams.microsoft.com/l/message/[CONVERSATION_ID]?tenantId=[TENANT_ID]&groupId=[GROUP_ID]&parentMessageId=[PARENT_MESSAGE_ID]&teamName=[TEAM_NAME]&channelName=[CHANNEL_ID]&createdTime=[PARENT_MESSAGE_ID]");
                    }}
                  >
                    Open Team Chat
                  </Button>
                  <Button
                    onClick={async () => {
                      await teamsSdk?.teamsJs.app.openLink("https://outlook.office.com/calendar/deeplink/compose?&subject=Sushi%20Training&location=Convention%20Center&startdt=2024-02-29T19%3A00%3A00&enddt=2024-03-01T00%3A00%3A05&body=Remember+to+bring+your+force!");
                    }}
                  >
                    Open Calendar invite
                  </Button>
                </div>
              }
              {this.state.selectedTab === "Mail" &&
                <div>
                  <Checkbox disabled label="mail" checked={teamsSdk?.teamsJs.mail.isSupported()} />
                  <Button disabled={!teamsSdk?.teamsJs.mail.isSupported()}
                    onClick={async () => {
                      try {
                        await teamsSdk?.teamsJs.mail.composeMail({
                          toRecipients: ["adelev@contoso.onmicrosoft.com"],
                          subject: "Test",
                          message: "Test",
                          type: teamsSdk?.teamsJs.mail.ComposeMailType.New,
                        });
                      } catch (e) {
                        console.log(e);
                        console.log(`GeoLocation error: ${e}`);
                      }
                    }}
                  >
                    Compose mail
                  </Button>
                </div>
              }
            </div>
          </div>
        </div>
      </FluentProvider>
    );
  }

  private _tabClickHandler = (e: SelectTabEvent, tab: SelectTabData): void => {
    this.setState({ selectedTab: tab.value });
  }
}
