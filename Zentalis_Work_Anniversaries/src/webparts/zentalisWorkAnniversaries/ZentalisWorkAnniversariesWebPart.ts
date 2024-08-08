import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'ZentalisWorkAnniversariesWebPartStrings';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IZentalisWorkAnniversariesWebPartProps {
  description: string;
  Title: string;
  AboutEmployee: string;
  Icon: string;
  Rightarrow: string;
  Leftarrow: string;
  AboutEmployeeLink: string;
}

export interface ISPList {
  displayName: string;
  employeeHireDate: string;
}

export default class ZentalisWorkAnniversariesWebPart extends BaseClientSideWebPart<IZentalisWorkAnniversariesWebPartProps> {

  private currentEmployeeIndex: number = 0;
  private employees: ISPList[] = [];
  private filteredEmployees: ISPList[] = [];
  private touchStartX: number = 0;
  private touchEndX: number = 0;
  private swipeThreshold: number = 50; // Minimum swipe distance to trigger a change

  public render(): void {
    this.context.msGraphClientFactory.getClient('3')
      .then((client: MSGraphClientV3): void => {
        client.api('https://graph.microsoft.com/v1.0/users?$select=displayName,employeeHireDate')
          .get((error, response: any, rawResponse?: any) => {
            if (error) {
              console.error(error);
              return;
            }

            if (!response || !response.value || response.value.length === 0) {
              console.warn('No data returned from Microsoft Graph API.');
              this.renderNoAnniversariesCard();
              return;
            }

            this.employees = response.value;
            this.filterEmployeesByCurrentQuarter();
            this.currentEmployeeIndex = 0;
            this.renderEmployeeCard();
          });
      });
  }

  private renderNoAnniversariesCard(): void {
    this.domElement.innerHTML = `
      <div class="anniversaries_Container">
        <div class="anniversaries_Title">
          <h3>${this.getCurrentQuarter()} ${this.properties.Title}</h3>
        </div>
        <div class="anniversaries_Lower_Celebrations">
          <p class="no_Anniversaries_Quarter">No work anniversaries this quarter.</p>
        </div>
      </div>
    `;
  }

  private filterEmployeesByCurrentQuarter(): void {
    const currentDate = new Date();
    const currentMonth = currentDate.getMonth();
    let quarterMonths: number[] = [];
    let quarterName: string = '';

    if (currentMonth >= 0 && currentMonth <= 2) {
      quarterMonths = [0, 1, 2]; // Q1: Jan, Feb, Mar
      quarterName = 'Q1';
    } else if (currentMonth >= 3 && currentMonth <= 5) {
      quarterMonths = [3, 4, 5]; // Q2: Apr, May, Jun
      quarterName = 'Q2';
    } else if (currentMonth >= 6 && currentMonth <= 8) {
      quarterMonths = [6, 7, 8]; // Q3: Jul, Aug, Sep
      quarterName = 'Q3';
    } else if (currentMonth >= 9 && currentMonth <= 11) {
      quarterMonths = [9, 10, 11]; // Q4: Oct, Nov, Dec
      quarterName = 'Q4';
    }

    this.filteredEmployees = this.employees.filter(employee => {
      const hireDate = new Date(employee.employeeHireDate);
      const yearsOfService = currentDate.getFullYear() - hireDate.getFullYear();
      const anniversaryPassedThisYear = 
        currentDate.getMonth() > hireDate.getMonth() ||
        (currentDate.getMonth() === hireDate.getMonth() && currentDate.getDate() >= hireDate.getDate());
      const yearsCompleted = anniversaryPassedThisYear ? yearsOfService : yearsOfService - 1;

      return quarterMonths.indexOf(hireDate.getMonth()) !== -1 && yearsCompleted > 0;
    });

    this.updateQuarterHeading(quarterName);
  }

  private updateQuarterHeading(quarterName: string): void {
    const headingElement = this.domElement.querySelector(`.anniversaries_Title h3`);
    if (headingElement) {
      headingElement.textContent = `${quarterName} Work Anniversaries`;
    }
  }

  private renderEmployeeCard(): void {
    if (this.filteredEmployees.length === 0) {
      this.renderNoAnniversariesCard();
      return;
    }

    const item = this.filteredEmployees[this.currentEmployeeIndex];
    const firstName = item.displayName.split(' ')[0];

    const hireDate = new Date(item.employeeHireDate);
    const currentDate = new Date();
    const yearsOfService = currentDate.getFullYear() - hireDate.getFullYear();
    const anniversaryPassedThisYear = 
      currentDate.getMonth() > hireDate.getMonth() ||
      (currentDate.getMonth() === hireDate.getMonth() && currentDate.getDate() >= hireDate.getDate());
    const yearsCompleted = anniversaryPassedThisYear ? yearsOfService : yearsOfService - 1;

    this.domElement.innerHTML = `
      <div class="anniversaries_Container">
        <div class="anniversaries_Title">
          <h3>${this.getCurrentQuarter()} ${this.properties.Title}</h3>
          <div class="anniversaries_Celebrations_Btn">
            <a href="#" id="prevEmployee"><img class="birthday_Celebrations_btn_Icon" src="${this.properties.Leftarrow}" alt="image"/></a>
            <a href="#" id="nextEmployee"><img class="birthday_Celebrations_btn_Icon" src="${this.properties.Rightarrow}" alt="image"/></a>
          </div>
        </div>
        <div class="anniversaries_Lower_Celebrations">
          <p class="anniversaries_Year">${yearsCompleted} ${yearsCompleted === 1 ? 'year' : 'years'}</p>
          <h2>${item.displayName}</h2>
          <p class="anniversaries_Text">Congrats ${firstName} for ${yearsCompleted} ${yearsCompleted === 1 ? 'year' : 'years'} at Zentalis!</p>
          <div class="anniversaries_CardsLink">
            <a href="${this.properties.AboutEmployeeLink}">${this.properties.AboutEmployee} ${firstName}</a>
            <img src="${this.properties.Icon}" alt="">
          </div>
        </div>
      </div>
    `;

    this.domElement.querySelector('#prevEmployee')?.addEventListener('click', this.showPreviousEmployee.bind(this));
    this.domElement.querySelector('#nextEmployee')?.addEventListener('click', this.showNextEmployee.bind(this));

    // Add touch event listeners for swipe functionality
    const cardContainer = this.domElement.querySelector('.anniversaries_Container');
    if (cardContainer) {
      cardContainer.addEventListener('touchstart', this.handleTouchStart.bind(this), false);
      cardContainer.addEventListener('touchmove', this.handleTouchMove.bind(this), false);
      cardContainer.addEventListener('touchend', this.handleTouchEnd.bind(this), false);
    }
  }

  private handleTouchStart(event: TouchEvent): void {
    this.touchStartX = event.changedTouches[0].clientX;
  }

  private handleTouchMove(event: TouchEvent): void {
    this.touchEndX = event.changedTouches[0].clientX;
  }

  private handleTouchEnd(event: TouchEvent): void {
    const swipeDistance = this.touchEndX - this.touchStartX;
    if (Math.abs(swipeDistance) > this.swipeThreshold) {
      if (swipeDistance > 0) {
        this.showPreviousEmployee(event);
      } else {
        this.showNextEmployee(event);
      }
    }
  }

  private showPreviousEmployee(event: Event): void {
    event.preventDefault();
    this.currentEmployeeIndex = (this.currentEmployeeIndex === 0) 
      ? this.filteredEmployees.length - 1 
      : this.currentEmployeeIndex - 1;
    this.renderEmployeeCard();
  }

  private showNextEmployee(event: Event): void {
    event.preventDefault();
    this.currentEmployeeIndex = (this.currentEmployeeIndex === this.filteredEmployees.length - 1) 
      ? 0 
      : this.currentEmployeeIndex + 1;
    this.renderEmployeeCard();
  }

  private getCurrentQuarter(): string {
    const currentMonth = new Date().getMonth();
    if (currentMonth >= 0 && currentMonth <= 2) return 'Q1';
    if (currentMonth >= 3 && currentMonth <= 5) return 'Q2';
    if (currentMonth >= 6 && currentMonth <= 8) return 'Q3';
    if (currentMonth >= 9 && currentMonth <= 11) return 'Q4';
    return '';
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
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }
          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
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
                PropertyPaneTextField('AboutEmployee', {
                  label: 'About Employee Text'
                }),
                PropertyPaneTextField('AboutEmployeeLink', {
                  label: 'About Employee Link'
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
