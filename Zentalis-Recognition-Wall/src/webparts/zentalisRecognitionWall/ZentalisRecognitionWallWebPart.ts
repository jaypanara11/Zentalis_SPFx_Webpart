import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ZentalisRecognitionWallWebPartStrings';
import ZentalisRecognitionWall from './components/ZentalisRecognitionWall';
import { IZentalisRecognitionWallProps } from './components/IZentalisRecognitionWallProps';

// Import the PeoplePicker and DateTimePicker
import {
  PropertyFieldPeoplePicker,
  PrincipalType
} from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import {
  PropertyFieldDateTimePicker,
  DateConvention,
  IDateTimeFieldValue
} from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';

import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

export interface IZentalisRecognitionWallWebPartProps {
  description: string;
  Header: string;
  Title: string;
  Description: string;
  LinkText: string;
  link: string;
  Icon: string;
  recognizedText1: string;
  Icon1: string;
  Description1: string;
  recognizedText2: string;
  Icon2: string;
  Description2: string;
  recognizedText3: string;
  Icon3: string;
  Description3: string;
  recognizedText4: string;
  Icon4: string;
  Description4: string;
  recognizedText5: string;
  Icon5: string;
  Description5: string;
  recognizedText6: string;
  Icon6: string;
  Description6: string;
  Date1: IDateTimeFieldValue;
  Date2: IDateTimeFieldValue;
  Date3: IDateTimeFieldValue;
  Date4: IDateTimeFieldValue;
  Date5: IDateTimeFieldValue;
  Date6: IDateTimeFieldValue;
  user: IPropertyFieldGroupOrPerson[];
  user1: IPropertyFieldGroupOrPerson[];
  user2: IPropertyFieldGroupOrPerson[];
  user3: IPropertyFieldGroupOrPerson[];
  user4: IPropertyFieldGroupOrPerson[];
  user5: IPropertyFieldGroupOrPerson[];
  user6: IPropertyFieldGroupOrPerson[];
  user7: IPropertyFieldGroupOrPerson[];
  user8: IPropertyFieldGroupOrPerson[];
  user9: IPropertyFieldGroupOrPerson[];
  user10: IPropertyFieldGroupOrPerson[];
  user11: IPropertyFieldGroupOrPerson[];
  showCard1: boolean;
  showCard2: boolean;
  showCard3: boolean;
  showCard4: boolean;
  showCard5: boolean;
  showCard6: boolean;
}

export default class ZentalisRecognitionWallWebPart extends BaseClientSideWebPart<IZentalisRecognitionWallWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IZentalisRecognitionWallProps> = React.createElement(
      ZentalisRecognitionWall,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        Header: this.properties.Header,
        Title: this.properties.Title,
        Description: this.properties.Description,
        LinkText: this.properties.LinkText,
        link: this.properties.link,
        Icon: this.properties.Icon,
        recognizedText1: this.properties.recognizedText1,
        Icon1: this.properties.Icon1,
        Description1: this.properties.Description1,
        recognizedText2: this.properties.recognizedText2,
        Icon2: this.properties.Icon2,
        Description2: this.properties.Description2,
        recognizedText3: this.properties.recognizedText3,
        Icon3: this.properties.Icon3,
        Description3: this.properties.Description3,
        recognizedText4: this.properties.recognizedText4,
        Icon4: this.properties.Icon4,
        Description4: this.properties.Description4,
        recognizedText5: this.properties.recognizedText5,
        Icon5: this.properties.Icon5,
        Description5: this.properties.Description5,
        recognizedText6: this.properties.recognizedText6,
        Icon6: this.properties.Icon6,
        Date1: this.properties.Date1,
        Date2: this.properties.Date2,
        Date3: this.properties.Date3,
        Date4: this.properties.Date4,
        Date5: this.properties.Date5,
        Date6: this.properties.Date6,
        Description6: this.properties.Description6,
        selectedUsers: this.properties.user,
        selectedUsers1: this.properties.user1,
        selectedUsers2: this.properties.user2,
        selectedUsers3: this.properties.user3,
        selectedUsers4: this.properties.user4,
        selectedUsers5: this.properties.user5,
        selectedUsers6: this.properties.user6,
        selectedUsers7: this.properties.user7,
        selectedUsers8: this.properties.user8,
        selectedUsers9: this.properties.user9,
        selectedUsers10: this.properties.user10,
        selectedUsers11: this.properties.user11,
        showCard1: this.properties.showCard1,
        showCard2: this.properties.showCard2,
        showCard3: this.properties.showCard3,
        showCard4: this.properties.showCard4,
        showCard5: this.properties.showCard5,
        showCard6: this.properties.showCard6
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
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
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
    const { semanticColors } = currentTheme;

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
            description: 'Navigation and Card 1'
          },
          groups: [
            {
              groupName: 'Recognition Wall Title property',
              groupFields: [
                PropertyPaneTextField('Header', {
                  label: 'Header'
                }),
                PropertyPaneTextField('Title', {
                  label: 'Title'
                }),
                PropertyPaneTextField('Description', {
                  label: 'Description'
                }),
                PropertyPaneTextField('LinkText', {
                  label: 'Navigation Text'
                }),
                PropertyPaneTextField('link', {
                  label: 'Navigation Link'
                }),
                PropertyPaneTextField('Icon', {
                  label: 'Navigation Icon'
                }),
              ]
            },
            {
              groupName: 'Card 1',
              groupFields: [
                PropertyPaneToggle('showCard1', {
                  label: 'Hide/Show',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyFieldPeoplePicker('user', {
                  label: 'Who Recognized',
                  initialData: this.properties.user,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  context: this.context as any,
                  properties: this.properties.user,
                  key: 'peopleFieldId'
                }),
                PropertyFieldPeoplePicker('user1', {
                  label: 'Recognized By',
                  initialData: this.properties.user1,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  context: this.context as any,
                  properties: this.properties.user1,
                  key: 'peopleFieldId'
                }),
                PropertyFieldDateTimePicker('Date1', {
                  label: 'Select the date',
                  initialDate: this.properties.Date1,
                  dateConvention: DateConvention.Date, // Use Date convention
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'dateTimeFieldId'
                }),
                PropertyPaneTextField('recognizedText1', {
                  label: 'Recognization Title'
                }),
                PropertyPaneTextField('Icon1', {
                  label: 'Card Icon'
                }),
                PropertyPaneTextField('Description1', {
                  label: 'Card Description'
                }),
              ]
            }
          ]
        },
        {
          header: {
            description: 'Card 2'
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneToggle('showCard2', {
                  label: 'Hide/Show',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyFieldPeoplePicker('user2', {
                  label: 'Who Recognized',
                  initialData: this.properties.user2,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  context: this.context as any,
                  properties: this.properties.user2,
                  key: 'peopleFieldId'
                }),
                PropertyFieldPeoplePicker('user3', {
                  label: 'Recognized By',
                  initialData: this.properties.user3,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  context: this.context as any,
                  properties: this.properties.user3,
                  key: 'peopleFieldId'
                }),
                PropertyFieldDateTimePicker('Date2', {
                  label: 'Select the date',
                  initialDate: this.properties.Date2,
                  dateConvention: DateConvention.Date, // Use Date convention
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'dateTimeFieldId'
                }),
                PropertyPaneTextField('recognizedText2', {
                  label: 'recognized for being'
                }),
                PropertyPaneTextField('Icon2', {
                  label: 'Recognition Icon'
                }),
                PropertyPaneTextField('Description2', {
                  label: 'User Description'
                }),
              ]
            }
          ]
        },
        {
          header: {
            description: 'Card 3'
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneToggle('showCard3', {
                  label: 'Hide/Show',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyFieldPeoplePicker('user4', {
                  label: 'Who Recognized',
                  initialData: this.properties.user4,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  context: this.context as any,
                  properties: this.properties.user4,
                  key: 'peopleFieldId'
                }),
                PropertyFieldPeoplePicker('user5', {
                  label: 'Recognized By',
                  initialData: this.properties.user5,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  context: this.context as any,
                  properties: this.properties.user5,
                  key: 'peopleFieldId'
                }),
                PropertyFieldDateTimePicker('Date3', {
                  label: 'Select the date',
                  initialDate: this.properties.Date3,
                  dateConvention: DateConvention.Date, // Use Date convention
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'dateTimeFieldId'
                }),
                PropertyPaneTextField('recognizedText3', {
                  label: 'recognized for being'
                }),
                PropertyPaneTextField('Icon3', {
                  label: 'Recognition Icon'
                }),
                PropertyPaneTextField('Description3', {
                  label: 'User Description'
                }),
              ]
            }
          ]
        },
        {
          header: {
            description: 'Card 4'
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneToggle('showCard4', {
                  label: 'Hide/Show',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyFieldPeoplePicker('user6', {
                  label: 'Who Recognized',
                  initialData: this.properties.user6,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  context: this.context as any,
                  properties: this.properties.user6,
                  key: 'peopleFieldId'
                }),
                PropertyFieldPeoplePicker('user7', {
                  label: 'Recognized By',
                  initialData: this.properties.user7,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  context: this.context as any,
                  properties: this.properties.user7,
                  key: 'peopleFieldId'
                }),
                PropertyFieldDateTimePicker('Date4', {
                  label: 'Select the date',
                  initialDate: this.properties.Date4,
                  dateConvention: DateConvention.Date, // Use Date convention
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'dateTimeFieldId'
                }),
                PropertyPaneTextField('recognizedText4', {
                  label: 'recognized for being'
                }),
                PropertyPaneTextField('Icon4', {
                  label: 'Recognition Icon'
                }),
                PropertyPaneTextField('Description4', {
                  label: 'User Description'
                }),
              ]
            }
          ]
        },
        {
          header: {
            description: 'Card 5'
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneToggle('showCard5', {
                  label: 'Hide/Show',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyFieldPeoplePicker('user8', {
                  label: 'Who Recognized',
                  initialData: this.properties.user8,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  context: this.context as any,
                  properties: this.properties.user8,
                  key: 'peopleFieldId'
                }),
                PropertyFieldPeoplePicker('user9', {
                  label: 'Recognized By',
                  initialData: this.properties.user9,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  context: this.context as any,
                  properties: this.properties.user9,
                  key: 'peopleFieldId'
                }),
                PropertyFieldDateTimePicker('Date5', {
                  label: 'Select the date',
                  initialDate: this.properties.Date5,
                  dateConvention: DateConvention.Date, // Use Date convention
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'dateTimeFieldId'
                }),
                PropertyPaneTextField('recognizedText5', {
                  label: 'recognized for being'
                }),
                PropertyPaneTextField('Icon5', {
                  label: 'Recognition Icon'
                }),
                PropertyPaneTextField('Description5', {
                  label: 'User Description'
                }),
              ]
            }
          ]
        },
        {
          header: {
            description: 'Card 6'
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneToggle('showCard6', {
                  label: 'Hide/Show',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyFieldPeoplePicker('user10', {
                  label: 'Who Recognized',
                  initialData: this.properties.user10,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  context: this.context as any,
                  properties: this.properties.user10,
                  key: 'peopleFieldId'
                }),
                PropertyFieldPeoplePicker('user11', {
                  label: 'Recognized By',
                  initialData: this.properties.user11,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  context: this.context as any,
                  properties: this.properties.user11,
                  key: 'peopleFieldId'
                }),
                PropertyFieldDateTimePicker('Date6', {
                  label: 'Select the date',
                  initialDate: this.properties.Date6,
                  dateConvention: DateConvention.Date, // Use Date convention
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'dateTimeFieldId'
                }),
                PropertyPaneTextField('recognizedText6', {
                  label: 'recognized for being'
                }),
                PropertyPaneTextField('Icon6', {
                  label: 'Recognition Icon'
                }),
                PropertyPaneTextField('Description6', {
                  label: 'User Description'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
