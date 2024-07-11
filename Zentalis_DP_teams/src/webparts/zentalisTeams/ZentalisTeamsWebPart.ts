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
//import styles from './ZentalisTeamsWebPart.module.scss';

export interface IZentalisTeamsWebPartProps {
  description: string;
  departmentFilter: string; 
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

  public render(): void {
    this.context.msGraphClientFactory
      .getClient('3')
      .then((client: MSGraphClientV3): void => {
        this._fetchDepartments(client).then(() => {
          this._fetchUsers(client).then(usersWithPhotos => {
            this.domElement.innerHTML = `
              <div>
                <div class="department_Team_Title">Department Teams</div>
                <div class="department_Team_Header">Meet the Teams</div>
                <div id="BindspListItems" class="Zentalis_Meet_Team">
                </div>
              </div>`;
            this._renderUsersByDepartment(usersWithPhotos);
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
      .select('id,displayName,jobTitle,department,employeeHireDate')
      .get();
      const users = response.value.map((user: any) => ({
        ...user,
        employeeHireDate: user.employeeHireDate
      }));
      return this._fetchUserPhotos(client, users);
    } catch (error) {
      console.error('Failed to fetch users:', error);
      return [];
    }
  }

  private async _fetchUserPhotos(client: MSGraphClientV3, users: ISPList[]): Promise<ISPList[]> {
    const usersWithPhotos = await Promise.all(users.map(async (user: ISPList) => {
      try {
        const photoResponse = await client.api(`https://graph.microsoft.com/v1.0/users/${user.id}/photo/$value`).responseType(ResponseType.BLOB).get();
        const photoUrl = URL.createObjectURL(photoResponse);
        return { ...user, photo: photoUrl };
      } catch (error) {
        console.error(`Failed to get photo for user ${user.displayName}:`, error);
        return { ...user, photo: undefined };
      }
    }));
    return usersWithPhotos;
  }

  private _renderUsersByDepartment(users: ISPList[]): void {
    const sortedUsers = users.sort((a, b) => new Date(a.employeeHireDate).getTime() - new Date(b.employeeHireDate).getTime());

    const filteredUsers = this._selectedDepartment ? sortedUsers.filter(user => user.department === this._selectedDepartment) : sortedUsers;
    const usersToDisplay = filteredUsers.slice(0, 8); 
    let html: string = `<div class="department_Container">`;

    for (let i = 0; i < usersToDisplay.length; i += 4) {
      html += `<div class="department_Row">`;

      for (let j = i; j < i + 4 && j < usersToDisplay.length; j++) {
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
                <div class="department_LearnAboutUser">Learn more about ${firstName}<img class="department_Nav_Icon" src="https://realitycraftprivatelimited.sharepoint.com/sites/DevJay/SiteAssets/Zentalis_Image/image.png" alt="Icon" /></div>
              </a>
            </div>
          </div>`;
      }

      html += `</div>`;
    }

    html += `</div>`;
    const listContainer: Element = this.domElement.querySelector('#BindspListItems')!;
    listContainer.innerHTML = html;
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
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneDropdown('departmentFilter', {
                  label: 'Filter by Department',
                  options: this._departmentOptions,
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
