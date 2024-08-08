import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'ZentalisTeamsWebPartStrings';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IZentalisTeamsWebPartProps {
  description: string;
  departmentFilter: string;
  Title: string;
  Header: string;
  Icon: string;
  Profile: string;
  CSSUrl: string;
}

export interface ISPList {
  displayName: string;
  id: string;
  photo?: string;
  jobTitle: string;
  department: string;
  employeeHireDate: string;
}

export default class ZentalisTeamsWebPart extends BaseClientSideWebPart<IZentalisTeamsWebPartProps> {

  private _selectedDepartment: string = '';
  private _departmentOptions: IPropertyPaneDropdownOption[] = [{ key: '', text: 'All' }];
  private _displayedUserCount: number = this._getInitialUserCount();

  public render(): void {
    const cssUrl = this.properties.CSSUrl || '';     
    if (cssUrl) {
      SPComponentLoader.loadCss(cssUrl);
    }

    this.context.msGraphClientFactory
      .getClient('3')
      .then((client: MSGraphClientV3): void => {
        this._fetchDepartments(client).then(() => {
          this._fetchUsers(client).then(usersWithPhotos => {
            this.domElement.innerHTML = `
              <div>
                <div class="department_Team_Title">${this.properties.Header}</div>
                <div class="department_Team_Header">${this.properties.Title}</div>
                <div id="BindspListItems" class="Zentalis_Meet_Team"></div>
                <button id="viewMoreButton" class="DP_Team_View_More" style="display: none;">View More</button>
              </div>`;
            this._renderUsersByDepartment(usersWithPhotos);

            const viewMoreButton = this.domElement.querySelector('#viewMoreButton') as HTMLElement;
            if (viewMoreButton) {
              viewMoreButton.addEventListener('click', () => this._toggleView(usersWithPhotos));
            }

            window.addEventListener('resize', () => this._onResize(usersWithPhotos));
          });
        });
      });
  }

  private async _fetchDepartments(client: MSGraphClientV3): Promise<void> {
    try {
      const response = await client.api('https://graph.microsoft.com/v1.0/users').select('department').get();
      const departments: string[] = Array.from(new Set(response.value.map((user: any) => user.department).filter((department: string | null) => department)));
      this._departmentOptions = [{ key: '', text: 'All' }, ...departments.map((department: string) => ({ key: department, text: department }))];
      this.context.propertyPane.refresh();
    } catch (error) {
      console.error('Failed to fetch departments:', error);
    }
  }

  private async _fetchUsers(client: MSGraphClientV3): Promise<ISPList[]> {
    try {
      const response = await client.api('https://graph.microsoft.com/v1.0/users')
        .select('id,displayName,jobTitle,department,employeeHireDate,accountEnabled')
        .get();
  
      const users = response.value
        .filter((user: any) => user.accountEnabled && !this._isJobTitleFiltered(user.jobTitle))
        .map((user: any) => ({
          ...user,
          employeeHireDate: user.employeeHireDate
        }));
  
      return this._fetchUserPhotos(client, users);
    } catch (error) {
      console.error('Failed to fetch users:', error);
      return [];
    }
  }
  
  private _isJobTitleFiltered(jobTitle: any): boolean {
    const titlesToFilter = ['Consultant', 'Intern'];
    const jobTitleStr = typeof jobTitle === 'string' ? jobTitle : '';
    return titlesToFilter.some(title => jobTitleStr.toLowerCase().includes(title.toLowerCase()));
  }
  
  

  private async _fetchUserPhotos(client: MSGraphClientV3, users: ISPList[]): Promise<ISPList[]> {
    const defaultPhotoUrl = this.properties.Profile || 'https://realitycraftprivatelimited.sharepoint.com/sites/DevJay/SiteAssets/Zentalis_Image/Profile_Image.png';
    const usersWithPhotos = await Promise.all(users.map(async (user: ISPList) => {
      try {
        if (user.id) {
          const photoResponse = await client.api(`https://graph.microsoft.com/v1.0/users/${user.id}/photo/$value`).responseType(ResponseType.BLOB).get();
          const photoUrl = URL.createObjectURL(photoResponse);
          return { ...user, photo: photoUrl };
        } else {
          return { ...user, photo: defaultPhotoUrl };
        }
      } catch (error) {
        console.error(`Failed to get photo for user ${user.displayName}:`, error);
        return { ...user, photo: defaultPhotoUrl };
      }
    }));

    return usersWithPhotos;
  }

  private _renderUsersByDepartment(users: ISPList[]): void {
    const sortedUsers = users.sort((a, b) => new Date(a.employeeHireDate).getTime() - new Date(b.employeeHireDate).getTime());

    const filteredUsers = this._selectedDepartment ? sortedUsers.filter(user => user.department === this._selectedDepartment) : sortedUsers;
    const usersToDisplay = filteredUsers.slice(0, this._displayedUserCount);

    let html: string = `<div class="department_Container">`;

    const usersPerRow = window.innerWidth <= 431 ? 3 : 4;
    for (let i = 0; i < usersToDisplay.length; i += usersPerRow) {
      html += `<div class="department_Row">`;

      for (let j = i; j < i + usersPerRow && j < usersToDisplay.length; j++) {
        const user = usersToDisplay[j];
        const firstName = user.displayName.split(' ')[0];
        html += `
          <div class="department_User_Item">
            ${user.photo ? `<img class="department_User_Image" src="${user.photo}" alt="${user.displayName}" />` : ''}
            <div class="department_User_Info">
              <div>
                <p class="department_UserName">${user.displayName}</p>
                <p class="department_UserJobTitle">${user.jobTitle}</p>
              </div>
              <a class="department_Ancher_link" href="#">
                <div class="department_LearnAboutUser">Learn more about ${firstName}<img class="department_Nav_Icon" src="${this.properties.Icon}" alt="Icon" /></div>
              </a>
            </div>
          </div>`;
      }

      html += `</div>`;
    }

    html += `</div>`;
    const listContainer: HTMLElement = this.domElement.querySelector('#BindspListItems') as HTMLElement;
    listContainer.innerHTML = html;

    const viewMoreButton = this.domElement.querySelector('#viewMoreButton') as HTMLElement;
    if (viewMoreButton) {
      if (filteredUsers.length > this._displayedUserCount) {
        viewMoreButton.style.display = 'block';
        viewMoreButton.textContent = this._displayedUserCount > this._getInitialUserCount() ? 'View Less' : 'View More';
      } else {
        viewMoreButton.style.display = 'none';
      }
    }
  }

  private _toggleView(users: ISPList[]): void {
    const initialUserCount = this._getInitialUserCount();
    if (window.innerWidth <= 431) {
      if (this._displayedUserCount === initialUserCount) {
        this._displayedUserCount = initialUserCount + 3;
      } else {
        this._displayedUserCount = initialUserCount;
      }
    } else {
      if (this._displayedUserCount === initialUserCount) {
        this._displayedUserCount = initialUserCount * 2;
      } else {
        this._displayedUserCount = initialUserCount;
      }
    }
    this._renderUsersByDepartment(users);
  }

  private _getInitialUserCount(): number {
    return window.innerWidth <= 431 ? 3 : 8;
  }

  private _onResize(users: ISPList[]): void {
    const initialUserCount = this._getInitialUserCount();
    if (this._displayedUserCount > initialUserCount) {
      this._displayedUserCount = initialUserCount;
    }
    this._renderUsersByDepartment(users);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      const savedDepartment = localStorage.getItem('selectedDepartment');
      if (savedDepartment) {
        this._selectedDepartment = savedDepartment;
      }
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
      const rootElement = this.domElement as HTMLElement;
      rootElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      rootElement.style.setProperty('--link', semanticColors.link || null);
      rootElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'departmentFilter') {
      this._selectedDepartment = newValue;
      localStorage.setItem('selectedDepartment', newValue);
      this.render();
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
                PropertyPaneDropdown('departmentFilter', {
                  label: 'Filter by Department',
                  options: this._departmentOptions,
                }),
                PropertyPaneTextField('Header', {
                  label: 'Team Header'
                }),
                PropertyPaneTextField('Title', {
                  label: 'Team Title'
                }),
                PropertyPaneTextField('Icon', {
                  label: 'Navigation Icon'
                }),
                PropertyPaneTextField('Profile', {
                  label: 'User Profile Image'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
