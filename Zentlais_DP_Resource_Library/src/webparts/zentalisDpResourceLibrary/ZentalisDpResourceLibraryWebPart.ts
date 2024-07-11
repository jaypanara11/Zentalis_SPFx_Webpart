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

export default class ZentalisDpResourceLibraryWebPart extends BaseClientSideWebPart<IZentalisDpResourceLibraryWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="zd_ResourceLibrary_Container">
        <div class="zd_ResourceLibrary_heading">
          <h5>Department Resources</h5>
          <h1>Resource Library</h1>
        </div>
        <div class="zd_ResourceLibrary_Navbar">
          <div class="tabs_nav">
            <a href="#Documents" class="nav_link">Documents</a>
            <a href="#Videos" class="nav_link">Videos</a>
            <a href="#Slides" class="nav_link">Slides</a>
            <a href="#TeamCalls" class="nav_link">Team Calls</a>
            <a href="#Policies" class="nav_link">Policies</a>
            <a href="#Apps" class="nav_link">Apps</a>
          </div>
        </div>
        <div class="zd_sorting_viewing">
          <div class="zd_sort">
            <div class="zd_sort_id">
              <p>Sort By: <span id="Az">Name:A to Z</span></p>
            </div>
            <div class="zd_sort_img"><a href=""><img src="" alt=""></a></div>
          </div>
          <div class="zd_viewAll">
            <div>
              <p id="viewAll" class="view_all_link">View All</p>
            </div>
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
      const activeSection = this.domElement.querySelector('.tab_content[style*="display: block;"]')!;
      const targetId = activeSection.id;
      this._showSection(targetId);
    });
  }

  private _showSection(sectionId: string): void {
    const sections = this.domElement.querySelectorAll('.tab_content');
    sections.forEach(section => {
      section.setAttribute('style', 'display: none;');
    });
    const activeSection = this.domElement.querySelector(`#${sectionId}`)!;
    activeSection.setAttribute('style', 'display: block;');

    // Update View All link text based on the active section
    const viewAllLink = this.domElement.querySelector('#viewAll')!;
    viewAllLink.textContent = `View All ${sectionId}`;
  }

  private _renderList(items: ISPList[]): void {
    let html: string = '<div class="zd_Main_Document"><div class="zd_Document_Left">';
    items.forEach((item: ISPList) => {
      html += `
        <div class="list_left">
          <a href="${item.EncodedAbsUrl}" target='_blank'>
            <div><img src="https://realitycraftprivatelimited.sharepoint.com/sites/DevJay/SiteAssets/Zentalis_Image/description.png" alt="">${item.File.Name}</div>
          </a>
        </div>
      `;
    });
    html += '</div></div>';
    const listContainer: Element = this.domElement.querySelector('#Documents')!;
    listContainer.innerHTML = html;
  }

  private _renderVideoList(items: ISPList[]): void {
    let html: string = '<div class="zd_video_upper"><div class="zd_video">';
    items.forEach((item: ISPList) => {
      html += `
        <div class="video_img">
          <a href="${item.EncodedAbsUrl}" target='_blank'>
            <img src="" alt="">
            <div class="video_text">
              <p>${item.File.Name}</p>
            </div>
          </a>
        </div>
      `;
    });
    html += '</div></div>';
    const listContainer: Element = this.domElement.querySelector('#Videos')!;
    listContainer.innerHTML = html;
  }

  private _renderImageList(items: ISPList[]): void {
    let html: string = '<div class="zd_Main_Slides"><div class="zd_slide_upper">';
    items.forEach((item: ISPList) => {
      html += `
        <div class="slide slide1">
          <div class="slide_img">
            <a href="${item.EncodedAbsUrl}" target='_blank'>
              <img src="${item.EncodedAbsUrl}" alt="">
              <div class="slide_text">
                <p>${item.File.Name}</p>
              </div>
            </a>
          </div>
        </div>
      `;
    });
    html += '</div></div>';
    const listContainer: Element = this.domElement.querySelector('#Slides')!;
    listContainer.innerHTML = html;
  }

  private _renderTeamCallsList(items: ISPList[]): void {
    let html: string = '<div class="zd_Main_Document"><div class="zd_Document_Left">';
    items.forEach((item: ISPList) => {
      html += `
        <div class="list_left">
          <a href="${item.EncodedAbsUrl}" target='_blank'>
            <div><img src="https://realitycraftprivatelimited.sharepoint.com/sites/DevJay/SiteAssets/Zentalis_Image/description.png" alt="">${item.File.Name}</div>
          </a>
        </div>
      `;
    });
    html += '</div></div>';
    const listContainer: Element = this.domElement.querySelector('#TeamCalls')!;
    listContainer.innerHTML = html;
  }

  private _renderPoliciesList(items: ISPList[]): void {
    let html: string = '<div class="zd_Main_Document"><div class="zd_Document_Left">';
    items.forEach((item: ISPList) => {
      html += `
        <div class="list_left">
          <a href="${item.EncodedAbsUrl}" target='_blank'>
            <div><img src="https://realitycraftprivatelimited.sharepoint.com/sites/DevJay/SiteAssets/Zentalis_Image/description.png" alt="">${item.File.Name}</div>
          </a>
        </div>
      `;
    });
    html += '</div></div>';
    const listContainer: Element = this.domElement.querySelector('#Policies')!;
    listContainer.innerHTML = html;
  }

  private _renderAppsList(items: ISPList[]): void {
    let html: string = '<div class="zd_Main_Document"><div class="zd_Document_Left">';
    items.forEach((item: ISPList) => {
      html += `
        <div class="list_left">
          <a href="${item.EncodedAbsUrl}" target='_blank'>
            <div><img src="https://realitycraftprivatelimited.sharepoint.com/sites/DevJay/SiteAssets/Zentalis_Image/description.png" alt="">${item.File.Name}</div>
          </a>
        </div>
      `;
    });
    html += '</div></div>';
    const listContainer: Element = this.domElement.querySelector('#Apps')!;
    listContainer.innerHTML = html;
  }

  private _getListData(listName: string): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('${listName}')/Items?$select=EncodedAbsUrl,*,File/Name&$expand=File`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error(`Error fetching data from list: ${listName}`);
        }
        return response.json();
      })
      .catch(error => {
        console.error(`Error fetching data from list: ${listName}`, error);
        return { value: [] };
      });
  }

  private _renderListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getListData('RL_Documents')
        .then((response) => {
          this._renderList(response.value);
        });
    }
  }

  private _renderVideoListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getListData('RL_Videos')
        .then((response) => {
          this._renderVideoList(response.value);
        });
    }
  }

  private _renderImageListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getListData('RL_Sliders')
        .then((response) => {
          this._renderImageList(response.value);
        });
    }
  }

  private _renderTeamCallsListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getListData('RL_TeamCalls')
        .then((response) => {
          this._renderTeamCallsList(response.value);
        });
    }
  }

  private _renderPoliciesListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getListData('RL_Policies')
        .then((response) => {
          this._renderPoliciesList(response.value);
        });
    }
  }

  private _renderAppsListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getListData('RL_Apps')
        .then((response) => {
          this._renderAppsList(response.value);
        });
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
