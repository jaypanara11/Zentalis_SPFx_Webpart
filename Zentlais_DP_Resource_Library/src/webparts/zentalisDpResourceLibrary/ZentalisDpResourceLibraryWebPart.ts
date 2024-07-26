import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'ZentalisDpResourceLibraryWebPartStrings';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

export interface IZentalisDpResourceLibraryWebPartProps {
  description: string;
  Icon: string;
  NavIcon: string;
  RLTitle: string;
  RLHeader: string;
  DLName1: string;
  DLName2: string;
  DLName3: string;
  DLName4: string;
  DLName5: string;
  DLName6: string;
  Section1: string;
  Section2: string;
  Section3: string;
  Section4: string;
  Section5: string;
  Section6: string;
  ViewAllLink1: string;
  ViewAllLink2: string;
  ViewAllLink3: string;
  ViewAllLink4: string;
  ViewAllLink5: string;
  ViewAllLink6: string;

}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  FileRef: string;
  EncodedAbsUrl: string;
  File: {
    Name: string;
  };
}

interface ISectionSortOrder {
  Documents: 'asc' | 'desc';
  Videos: 'asc' | 'desc';
  Slides: 'asc' | 'desc';
  TeamCalls: 'asc' | 'desc';
  Policies: 'asc' | 'desc';
  Apps: 'asc' | 'desc';
}

export default class ZentalisDpResourceLibraryWebPart extends BaseClientSideWebPart<IZentalisDpResourceLibraryWebPartProps> {

  private currentSortOrder: ISectionSortOrder = {
    Documents: 'asc',
    Videos: 'asc',
    Slides: 'asc',
    TeamCalls: 'asc',
    Policies: 'asc',
    Apps: 'asc'
  };

  public render(): void {
    this.domElement.innerHTML = `
      <div class="zd_ResourceLibrary_Container">
        <div class="zd_ResourceLibrary_heading">
          <h5>${this.properties.RLTitle}</h5>
          <h1>${this.properties.RLHeader}</h1>
        </div>
        <div class="zd_ResourceLibrary_Navbar">
          <div class="tabs_nav">
            <a href="#Documents" class="nav_link">${this.properties.Section1}</a>
            <a href="#Videos" class="nav_link">${this.properties.Section2}</a>
            <a href="#Slides" class="nav_link">${this.properties.Section3}</a>
            <a href="#TeamCalls" class="nav_link">${this.properties.Section4}</a>
            <a href="#Policies" class="nav_link">${this.properties.Section5}</a>
            <a href="#Apps" class="nav_link">${this.properties.Section6}</a>
          </div>
          <div class="dropdown">
            <button id="dropdownButton" class="dropbtn" onclick="toggleDropdown()">Menu</button>
            <div id="dropdownContent" class="dropdown_content">
              <a href="#Documents" class="nav_link" onclick="selectOption('Documents')">${this.properties.Section1}</a>
              <a href="#Videos" class="nav_link" onclick="selectOption('Videos')">${this.properties.Section2}</a>
              <a href="#Slides" class="nav_link" onclick="selectOption('Slides')">${this.properties.Section3}</a>
              <a href="#TeamCalls" class="nav_link" onclick="selectOption('Team Calls')">${this.properties.Section4}</a>
              <a href="#Policies" class="nav_link" onclick="selectOption('Policies')">${this.properties.Section5}</a>
              <a href="#Apps" class="nav_link" onclick="selectOption('Apps')">${this.properties.Section6}</a>
            </div>
          </div>
        </div>
        <div class="zd_sorting_viewing">
          <div class="zd_sort">
            <div class="zd_sort_id">
              <p>Sort By: <span id="sortName" style="cursor: pointer;">Name: A to Z</span></p>
            </div>
            <div class="zd_sort_img"></div>
          </div>
          <div class="zd_viewAll">           
              <a id="viewAll" class="view_all_link" href="#">View All</a><img src="${this.properties.NavIcon}" alt="Icon"/>           
          </div>
        </div>
        <div class="tabs">
          <div id="Documents" class="tab_content" style="display: block;"></div>
          <div id="Videos" class="tab_content" style="display: none;"></div>
          <div id="Slides" class="tab_content" style="display: none;"></div>
          <div id="TeamCalls" class="tab_content" style="display: none;"></div>
          <div id="Policies" class="tab_content" style="display: none;"></div>
          <div id="Apps" class="tab_content" style="display: none;"></div>
        </div>
        <div class="Bottom_ViewAll">
          <button id="BottomviewAll" class="Bottom_view_all_link">View All</button>
        </div>

        <div id="DocumentDetails"></div>
      </div>
    `;

    this._addEventListeners();
    this._showSection('Documents');
    this._renderListAsync();
    this._renderVideoListAsync();
    this._renderImageListAsync();
    this._renderTeamCallsListAsync();
    this._renderPoliciesListAsync();
    this._renderAppsListAsync();
  }

  private _addEventListeners(): void {
    const navLinks = this.domElement.querySelectorAll('.nav_link');
    navLinks.forEach(link => {
      link.addEventListener('click', (event) => {
        event.preventDefault();
        const targetId = (event.target as HTMLAnchorElement).getAttribute('href')!.substring(1);
        this._showSection(targetId);
      });
    });

    const viewAllLink = this.domElement.querySelector('#viewAll')!;
    viewAllLink.addEventListener('click', (event) => {
      event.preventDefault();
      const activeSection = this.domElement.querySelector('.tab_content[style*="display: block;"]')!;
      const targetId = activeSection.id;
      const sectionUrls: { [key: string]: string; } = {
        'Documents': this.properties.ViewAllLink1,
        'Videos': this.properties.ViewAllLink2,
        'Slides': this.properties.ViewAllLink3,
        'TeamCalls': this.properties.ViewAllLink4,
        'Policies': this.properties.ViewAllLink5,
        'Apps': this.properties.ViewAllLink6,
      };
      const viewAllUrl = sectionUrls[targetId];
      this._openUrl(viewAllUrl);
    });

    const bottomViewAllButton = this.domElement.querySelector('#BottomviewAll')!;
    if (bottomViewAllButton) {
      bottomViewAllButton.addEventListener('click', (event) => {
        event.preventDefault();
        this._handleBottomViewAllClick();
      });
    }
  

    const sortSpan = this.domElement.querySelector('#sortName')!;
    sortSpan.addEventListener('click', () => {
      const activeSection = this.domElement.querySelector('.tab_content[style*="display: block;"]')!;
      const sectionId = activeSection.id as keyof ISectionSortOrder;
      this.currentSortOrder[sectionId] = this.currentSortOrder[sectionId] === 'asc' ? 'desc' : 'asc';
      sortSpan.textContent = `Name: ${this.currentSortOrder[sectionId] === 'asc' ? 'A to Z' : 'Z to A'}`;
      this._renderSection(sectionId);

    });
    const dropdownButton = document.getElementById('dropdownButton');
    if (dropdownButton) {
      dropdownButton.addEventListener('click', () => this.toggleDropdown());
    }
  }
  private _handleBottomViewAllClick(): void {
    const activeSection = this.domElement.querySelector('.tab_content[style*="display: block;"]')!;
    const targetId = activeSection.id;
    this._renderAllItems(targetId as keyof ISectionSortOrder);
  }
  

     private toggleDropdown(): void {
    const dropdown = this.domElement.querySelector('.dropdown') as HTMLElement;
    dropdown.classList.toggle('show');
  }


  private _showSection(sectionId: string): void {
    const sections = this.domElement.querySelectorAll('.tab_content');
    sections.forEach(section => {
      section.setAttribute('style', 'display: none;');
    });
    const activeSection = this.domElement.querySelector(`#${sectionId}`) as HTMLElement;
    activeSection.style.display = 'block';
  
    const navLinks = this.domElement.querySelectorAll('.nav_link');
    navLinks.forEach(link => {
      const navLink = link as HTMLAnchorElement;
      if (navLink.getAttribute('href') === `#${sectionId}`) {
        navLink.style.backgroundColor = '#3FA8F4'; 
        navLink.style.padding = '10px 16px';
        navLink.style.borderTopLeftRadius = '12px'; 
        navLink.style.borderTopRightRadius = '12px'; 
        navLink.style.color = '#FFFFFF';
      } else {
        navLink.style.backgroundColor = '';
        navLink.style.padding = '';
        navLink.style.borderTopLeftRadius = ''; 
        navLink.style.borderTopRightRadius = ''; 
        navLink.style.color = '';
      }
    });
  
    const viewAllLink = this.domElement.querySelector('#viewAll') as HTMLAnchorElement;
    const sectionUrls: { [key: string]: string; } = {
      'Documents': this.properties.ViewAllLink1,
      'Videos': this.properties.ViewAllLink2,
      'Slides': this.properties.ViewAllLink3,
      'TeamCalls': this.properties.ViewAllLink4,
      'Policies': this.properties.ViewAllLink5,
      'Apps': this.properties.ViewAllLink6,
    };
    viewAllLink.setAttribute('href', sectionUrls[sectionId]);
    viewAllLink.textContent = `View All `;
  }

  private _sortItems(items: ISPList[], sectionId: keyof ISectionSortOrder): ISPList[] {
    return items.sort((a, b) => {
      const nameA = a.File.Name.toLowerCase();
      const nameB = b.File.Name.toLowerCase();

      if (this.currentSortOrder[sectionId] === 'asc') {
        return nameA < nameB ? -1 : nameA > nameB ? 1 : 0;
      } else {
        return nameA > nameB ? -1 : nameA < nameB ? 1 : 0;
      }
    });
  }

  private _renderSection(sectionId: keyof ISectionSortOrder): void {
    switch (sectionId) {
      case 'Documents':
        this._renderListAsync();
        break;
      case 'Videos':
        this._renderVideoListAsync();
        break;
      case 'Slides':
        this._renderImageListAsync();
        break;
      case 'TeamCalls':
        this._renderTeamCallsListAsync();
        break;
      case 'Policies':
        this._renderPoliciesListAsync();
        break;
      case 'Apps':
        this._renderAppsListAsync();
        break;
      default:
        break;
    }
  }

  private _renderAllItems(sectionId: keyof ISectionSortOrder): void {
    switch (sectionId) {
      case 'Documents':
        this._renderFullListAsync();
        break;
      case 'Videos':
        this._renderFullVideoListAsync();
        break;
      case 'Slides':
        this._renderFullImageListAsync();
        break;
      case 'TeamCalls':
        this._renderFullTeamCallsListAsync();
        break;
      case 'Policies':
        this._renderFullPoliciesListAsync();
        break;
      case 'Apps':
        this._renderFullAppsListAsync();
        break;
      default:
        break;
    }
  }

  private _renderList(items: ISPList[]): void {
    items = this._sortItems(items, 'Documents')
    let html: string = '<div class="zd_Main_Document"><div class="zd_Document_Left">';
    items.forEach((item: ISPList) => {
      html += `
        <div class="list_left">
          <a href="${item.EncodedAbsUrl}" target='_blank'>
            <div><img src="${this.properties.Icon}" alt="">${item.File.Name}</div>
          </a>
        </div>
      `;
    });
    html += '</div></div>';
    const listContainer: Element = this.domElement.querySelector('#Documents')!;
    listContainer.innerHTML = html;
  }

  private _renderVideoList(items: ISPList[]): void {
    items = this._sortItems(items, 'Videos')
    let html: string = '<div class="zd_Main_Slides"><div class="zd_slide_upper">';
    items.forEach((item: ISPList) => {
      html += `
        <div class="slide">
          <a href="${item.EncodedAbsUrl}" target='_blank'>
            <div class="Video_img">
              <video controls>
                <source src="${item.EncodedAbsUrl}" type="video/mp4">
              </video>
            </div>
            <div class="slide_text">${item.File.Name}</div>
          </a>
        </div>
      `;
    });
    html += '</div></div>';
    const listContainer: Element = this.domElement.querySelector('#Videos')!;
    listContainer.innerHTML = html;
  }

  private _renderImageList(items: ISPList[]): void {
    items = this._sortItems(items, 'Slides')
    let html: string = '<div class="zd_Main_Slides"><div class="zd_slide_upper">';
    items.forEach((item: ISPList) => {
      html += `
        <div class="slide">
          <a href="${item.EncodedAbsUrl}" target='_blank'>
            <div class="slide_img">            
              <img src="${item.EncodedAbsUrl}" alt="">                         
            </div>
            <div class="slide_text">${item.File.Name}</div>
          </a>
        </div>
      `;
    });
    html += '</div></div>';
    const listContainer: Element = this.domElement.querySelector('#Slides')!;
    listContainer.innerHTML = html;
  }

  private _renderTeamCallsList(items: ISPList[]): void {
    items = this._sortItems(items, 'TeamCalls')
    let html: string = '<div class="zd_Main_Document"><div class="zd_Document_Left">';
    items.forEach((item: ISPList) => {
      html += `
        <div class="list_left">
          <a href="${item.EncodedAbsUrl}" target='_blank'>
            <div><img src="${this.properties.Icon}" alt="">${item.File.Name}</div>
          </a>
        </div>
      `;
    });
    html += '</div></div>';
    const listContainer: Element = this.domElement.querySelector('#TeamCalls')!;
    listContainer.innerHTML = html;
  }

  private _renderPoliciesList(items: ISPList[]): void {
    items = this._sortItems(items, 'Policies')
    let html: string = '<div class="zd_Main_Document"><div class="zd_Document_Left">';
    items.forEach((item: ISPList) => {
      html += `
        <div class="list_left">
          <a href="${item.EncodedAbsUrl}" target='_blank'>
            <div><img src="${this.properties.Icon}" alt="">${item.File.Name}</div>
          </a>
        </div>
      `;
    });
    html += '</div></div>';
    const listContainer: Element = this.domElement.querySelector('#Policies')!;
    listContainer.innerHTML = html;
  }

  private _renderAppsList(items: ISPList[]): void {
    items = this._sortItems(items, 'Apps')
    let html: string = '<div class="zd_Main_Slides"><div class="zd_slide_upper">';
    items.forEach((item: ISPList) => {
      html += `
        <div class="slide">
          <a href="${item.EncodedAbsUrl}" target='_blank'>
            <div class="slide_img">            
              <div class="App_Box"></div>                          
            </div>
            <div class="slide_text">${item.File.Name}</div>
          </a>
        </div>
      `;
    });
    html += '</div></div>';
    const listContainer: Element = this.domElement.querySelector('#Apps')!;
    listContainer.innerHTML = html;
  }

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('"+this.properties.DLName1+"')/Items?$select=EncodedAbsUrl,*,File/Name&$expand=File", SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error('Error fetching list data');
        }
        return response.json();
      })
      .catch(error => {
        console.error('Error fetching list data', error);
        return { value: [] };
      });
  }

  private _getVideoListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('"+this.properties.DLName2+"')/Items?$select=EncodedAbsUrl,*,File/Name&$expand=File", SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error('Error fetching video list data');
        }
        return response.json();
      })
      .catch(error => {
        console.error('Error fetching video list data', error);
        return { value: [] };
      });
  }

  private _getImageListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('"+this.properties.DLName3+"')/Items?$select=EncodedAbsUrl,*,File/Name&$expand=File", SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error('Error fetching image list data');
        }
        return response.json();
      })
      .catch(error => {
        console.error('Error fetching image list data', error);
        return { value: [] };
      });
  }

  private _getTeamCallsData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('"+this.properties.DLName4+"')/Items?$select=EncodedAbsUrl,*,File/Name&$expand=File", SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error('Error fetching team calls data');
        }
        return response.json();
      })
      .catch(error => {
        console.error('Error fetching team calls data', error);
        return { value: [] };
      });
  }

  private _getPoliciesData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('"+this.properties.DLName5+"')/Items?$select=EncodedAbsUrl,*,File/Name&$expand=File", SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error('Error fetching policies data');
        }
        return response.json();
      })
      .catch(error => {
        console.error('Error fetching policies data', error);
        return { value: [] };
      });
  }

  private _getAppsData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('"+this.properties.DLName6+"')/Items?$select=EncodedAbsUrl,*,File/Name&$expand=File", SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error('Error fetching apps data');
        }
        return response.json();
      })
      .catch(error => {
        console.error('Error fetching apps data', error);
        return { value: [] };
      });
  }

  private _renderListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          this._renderList(response.value.slice(0, 12)); // Initial render limited to 12 items
        });
    }
  }
  
  private _renderFullListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          this._renderList(response.value); // Render all items
        });
    }
  }
  
  private _renderVideoListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getVideoListData()
        .then((response) => {
          this._renderVideoList(response.value.slice(0, 8)); // Initial render limited to 8 items
        });
    }
  }
  
  private _renderFullVideoListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getVideoListData()
        .then((response) => {
          this._renderVideoList(response.value); // Render all items
        });
    }
  }
  
  private _renderImageListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getImageListData()
        .then((response) => {
          this._renderImageList(response.value.slice(0, 8)); // Initial render limited to 8 items
        });
    }
  }
  
  private _renderFullImageListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getImageListData()
        .then((response) => {
          this._renderImageList(response.value); // Render all items
        });
    }
  }
  
  private _renderTeamCallsListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getTeamCallsData()
        .then((response) => {
          this._renderTeamCallsList(response.value.slice(0, 12)); // Initial render limited to 12 items
        });
    }
  }
  
  private _renderFullTeamCallsListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getTeamCallsData()
        .then((response) => {
          this._renderTeamCallsList(response.value); // Render all items
        });
    }
  }
  
  private _renderPoliciesListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getPoliciesData()
        .then((response) => {
          this._renderPoliciesList(response.value.slice(0, 12)); // Initial render limited to 12 items
        });
    }
  }
  
  private _renderFullPoliciesListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getPoliciesData()
        .then((response) => {
          this._renderPoliciesList(response.value); // Render all items
        });
    }
  }
  
  private _renderAppsListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getAppsData()
        .then((response) => {
          this._renderAppsList(response.value.slice(0, 8)); // Initial render limited to 8 items
        });
    }
  }
  
  private _renderFullAppsListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getAppsData()
        .then((response) => {
          this._renderAppsList(response.value); // Render all items
        });
    }
  }
  private _openUrl(url: string): void {
    window.open(url, '_blank');
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
              groupName: 'Top Navigation',
              groupFields: [
                PropertyPaneTextField('RLTitle', {
                  label: 'Resource Library Title'
                }),
                PropertyPaneTextField('RLHeader', {
                  label: 'Resource Library Header'
                }),
                PropertyPaneTextField('Icon', {
                  label: 'Document Icon'
                }),
                PropertyPaneTextField('NavIcon', {
                  label: 'Navigation Arrow Icon'
                }),
              ]
            },
            {
              groupName: 'Document Section',
              groupFields: [
                PropertyPaneTextField('DLName1', {
                  label: 'Document Library Name'
                }),
                PropertyPaneTextField('Section1', {
                  label: 'Section Name'
                }),
                PropertyPaneTextField('ViewAllLink1', {
                  label: 'View All Link'
                }),
              ]
            },
            {
              groupName: 'Videos Section',
              groupFields: [
                PropertyPaneTextField('DLName2', {
                  label: 'Document Library Name'
                }),
                PropertyPaneTextField('Section2', {
                  label: 'Section Name'
                }),
                PropertyPaneTextField('ViewAllLink2', {
                  label: 'View All Link'
                }),
              ]
            },
            {
              groupName: 'Slides Section',
              groupFields: [
                PropertyPaneTextField('DLName3', {
                  label: 'Document Library Name'
                }),
                PropertyPaneTextField('Section3', {
                  label: 'Section Name'
                }),
                PropertyPaneTextField('ViewAllLink3', {
                  label: 'View All Link'
                }),
              ]
            },
            {
              groupName: 'Team Calls Section',
              groupFields: [
                PropertyPaneTextField('DLName4', {
                  label: 'Document Library Name'
                }),
                PropertyPaneTextField('Section4', {
                  label: 'Section Name'
                }),
                PropertyPaneTextField('ViewAllLink4', {
                  label: 'View All Link'
                }),
              ]
            },
            {
              groupName: 'Policies Section',
              groupFields: [
                PropertyPaneTextField('DLName5', {
                  label: 'Document Library Name'
                }),
                PropertyPaneTextField('Section5', {
                  label: 'Section Name'
                }),
                PropertyPaneTextField('ViewAllLink5', {
                  label: 'View All Link'
                }),
              ]
            },
            {
              groupName: 'Apps Section',
              groupFields: [
                PropertyPaneTextField('DLName6', {
                  label: 'Document Library Name'
                }),
                PropertyPaneTextField('Section6', {
                  label: 'Section Name'
                }),
                PropertyPaneTextField('ViewAllLink6', {
                  label: 'View All Link'
                }),
              ]
            },
          ]
        }
      ]
    };
  }
}
