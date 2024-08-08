import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'ZentalisBirthdayWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IZentalisBirthdayWebPartProps {
  description: string;
  ListName: string;
  Title: string;
  AboutEmployee: string;
  Icon: string;
  Rightarrow: string;
  Leftarrow: string;
  AboutEmployeeLink: string;
}

export interface ISPList {
  EmployeeName: { Title: string };
  EmployeeBirthdate: string;
}

export default class ZentalisBirthdayWebPart extends BaseClientSideWebPart<IZentalisBirthdayWebPartProps> {

  private currentEmployeeIndex: number = 0;
  private employees: ISPList[] = [];
  private touchStartX: number = 0;
  private touchEndX: number = 0;
  private swipeThreshold: number = 50; // Minimum swipe distance to trigger a change

  public render(): void {
    this.fetchData()
      .then(() => {
        this.currentEmployeeIndex = 0; // Start with the first employee
        this.renderEmployeeCard();
      });
  }

  private fetchData(): Promise<void> {
    const url: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.ListName}')/items?$select=EmployeeName/Title,EmployeeBirthdate&$expand=EmployeeName`;

    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((data: any) => {
        this.employees = data.value.filter((employee: ISPList) => this.isCurrentQuarter(new Date(employee.EmployeeBirthdate)));
      })
      .catch(error => {
        console.error(error);
      });
  }

  private isCurrentQuarter(date: Date): boolean {
    const currentMonth = new Date().getMonth(); // 0-11
    const currentQuarter = Math.floor(currentMonth / 3) + 1;

    const employeeMonth = date.getMonth(); // 0-11
    const employeeQuarter = Math.floor(employeeMonth / 3) + 1;

    return currentQuarter === employeeQuarter;
  }

  private renderEmployeeCard(): void {
    const currentQuarter = this.getCurrentQuarter();
    const quarterTitle = `Q${currentQuarter} ${this.properties.Title}`;

    if (this.employees.length === 0) {
      this.domElement.innerHTML = `
        <div class="birthday_Container">
          <div class="birthdays_Title">
            <h3>${quarterTitle}</h3>
          </div>
          <div class="birthday_Lower_Celebrations">
            <div class="no_Birthday_Quarter">No Birthday in this quarter.</div>
          </div>
        </div>
      `;
      return;
    }

    const item = this.employees[this.currentEmployeeIndex];
    const employeeName = item.EmployeeName.Title;
    const firstName = employeeName.split(' ')[0];
    const birthDate = new Date(item.EmployeeBirthdate);
    const formattedDate = `${birthDate.toLocaleString('default', { month: 'long' })} ${birthDate.getDate()}<sup>${this.getDaySuffix(birthDate.getDate())}</sup>`;

    this.domElement.innerHTML = `
        <div class="birthday_Container">
          <div class="birthdays_Title">
            <h3>${quarterTitle}</h3>
            <div class="birthday_Celebrations_btn">
              <a href="#" id="prevEmployee"><img class="birthday_Celebrations_btn_Icon" src="${this.properties.Leftarrow}" alt="Previous"/></a>
              <a href="#" id="nextEmployee"><img class="birthday_Celebrations_btn_Icon" src="${this.properties.Rightarrow}" alt="Next"/></a>
            </div>
          </div>
          <div class="birthday_Lower_Celebrations">
            <p class="birthday_Year">${formattedDate}</p>
            <h2>${employeeName}</h2>
            <div class="birthday_Cards_Link">
              <a href="${this.properties.AboutEmployeeLink}">${this.properties.AboutEmployee} ${firstName}</a>
              <img src="${this.properties.Icon}" />
            </div>
          </div>
        </div>
    `;

    // Attach event listeners for prev and next buttons
    const prevButton = this.domElement.querySelector('#prevEmployee');
    const nextButton = this.domElement.querySelector('#nextEmployee');
    if (prevButton) {
      prevButton.addEventListener('click', this.showPreviousEmployee.bind(this));
    }
    if (nextButton) {
      nextButton.addEventListener('click', this.showNextEmployee.bind(this));
    }

    // Add touch event listeners for swipe functionality
    const cardContainer = this.domElement.querySelector('.birthday_Container');
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
    if (this.employees.length > 0) {
      this.currentEmployeeIndex = (this.currentEmployeeIndex === 0) 
        ? this.employees.length - 1 
        : this.currentEmployeeIndex - 1;
      this.renderEmployeeCard();
    }
  }

  private showNextEmployee(event: Event): void {
    event.preventDefault();
    if (this.employees.length > 0) {
      this.currentEmployeeIndex = (this.currentEmployeeIndex === this.employees.length - 1) 
        ? 0 
        : this.currentEmployeeIndex + 1;
      this.renderEmployeeCard();
    }
  }

  private getCurrentQuarter(): number {
    const currentMonth = new Date().getMonth(); // 0-11
    return Math.floor(currentMonth / 3) + 1;
  }

  private getDaySuffix(day: number): string {
    if (day > 3 && day < 21) return 'th'; // special case for 11th-13th
    switch (day % 10) {
      case 1: return 'st';
      case 2: return 'nd';
      case 3: return 'rd';
      default: return 'th';
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
            description: ''
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('ListName', {
                  label: 'List Name'
                }),
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
