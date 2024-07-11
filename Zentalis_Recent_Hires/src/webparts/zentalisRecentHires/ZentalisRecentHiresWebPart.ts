import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

//import styles from './ZentalisRecentHiresWebPart.module.scss';
import * as strings from 'ZentalisRecentHiresWebPartStrings';
import { MSGraphClientV3 } from '@microsoft/sp-http';

// import { SPComponentLoader } from '@microsoft/sp-loader'; 
// const cssUrl = `https://realitycraftprivatelimited.sharepoint.com/sites/DevJay/SiteAssets/Zentalis.css`;
// SPComponentLoader.loadCss(cssUrl);

export interface IZentalisRecentHiresWebPartProps {
  description: string;
  Title: string;
  WelcomeText: string;
  aboutEmployee: string;
  Icon: string;
  Rightarrow: string;
  Leftarrow: string;
}

export interface ISPList {
  displayName: string;
  employeeHireDate: string;
  jobTitle: string;
  city: string;
  state: string;
}
export default class ZentalisRecentHiresWebPart extends BaseClientSideWebPart<IZentalisRecentHiresWebPartProps> {

  private currentEmployeeIndex: number = 0;
  private employees: ISPList[] = [];

  public render(): void {
    this.context.msGraphClientFactory.getClient('3')
      .then((client: MSGraphClientV3): void => {
        client.api('https://graph.microsoft.com/v1.0/users?$select=displayName,employeeHireDate,jobTitle,city,state')
          .top(10)
          .get((error, response: any, rawResponse?: any) => {
            if (error) {
              console.error(error);
              return;
            }

            if (!response || !response.value || response.value.length === 0) {
              console.warn('No data returned from Microsoft Graph API.');
              return;
            }

            this.employees = response.value;
            this.currentEmployeeIndex = 0; // Start with the first employee
            this.renderEmployeeCard();
          });
      });
  }

  private renderEmployeeCard(): void {
    const item = this.employees[this.currentEmployeeIndex];
    const firstName = item.displayName.split(' ')[0];

    this.domElement.innerHTML = `
      <div class="recentHireContainer">
        <div class="recent_Hiring">
          <div class="recent_Hiring_Title">
            <h3>${this.properties.Title}</h3>
            <div class="recent_Hiring_Celebrations_Btn">
              <a href="#" id="prevEmployee"><img class="birthday_Celebrations_btn_Icon" src="${this.properties.Leftarrow}" alt="image"/></a>
              <a href="#" id="nextEmployee"><img class="birthday_Celebrations_btn_Icon" src="${this.properties.Rightarrow}" alt="image"/></a>
            </div>
          </div>
          <div class="recent_Hiring_Lower_Celebrations">
            <h2>${this.properties.WelcomeText},<br> ${item.displayName}!</h2>
            <p>${item.jobTitle} from ${item.city},<br> ${item.state}!</p>
            <div class="recent_Hiring_Cardslink">
              <a href="#">${this.properties.aboutEmployee} ${firstName}</a>
              <img src="${this.properties.Icon}" alt="">
            </div>
          </div>
        </div>
      </div>`;

    // Attach event listeners for prev and next buttons
    this.domElement.querySelector('#prevEmployee')?.addEventListener('click', this.showPreviousEmployee.bind(this));
    this.domElement.querySelector('#nextEmployee')?.addEventListener('click', this.showNextEmployee.bind(this));
  }

  private showPreviousEmployee(event: Event): void {
    event.preventDefault();
    if (this.currentEmployeeIndex > 0) {
      this.currentEmployeeIndex--;
      this.renderEmployeeCard();
    }
  }

  private showNextEmployee(event: Event): void {
    event.preventDefault();
    if (this.currentEmployeeIndex < this.employees.length - 1) {
      this.currentEmployeeIndex++;
      this.renderEmployeeCard();
    }
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      console.log(message);
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

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
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
                PropertyPaneTextField('Title', {
                  label: 'Title'
                }),
                PropertyPaneTextField('WelcomeText', {
                  label: 'Welcome Text'
                }),
                PropertyPaneTextField('aboutEmployee', {
                  label: 'About Employee'
                }),
                PropertyPaneTextField('Icon', {
                  label: 'Navigation Icon'
                }),
                PropertyPaneTextField('Leftarrow', {
                  label: 'Left Arrow'
                }),
                PropertyPaneTextField('Rightarrow', {
                  label: 'Right Arrow'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
