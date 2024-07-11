import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader'; 

const cssUrl = `https://zenopharma.sharepoint.com/sites/ZenSPDev/SiteAssets/ZentalisDepartment.css`;
SPComponentLoader.loadCss(cssUrl);

import * as strings from 'ZentalisDepBannerWebPartStrings';

export interface IZentalisDepBannerWebPartProps {
  description: string;
  ListName: string;
  WelcomeText: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  SubTitle: string;
  TilesBackgroundColor: string;
  TitleColor: string;
  SubTitleColor: string;
  Links: {
    Url: string;
    Description: string;
  };
  SectionPosition: string;
  IsActive: boolean;
  OrderID: number;
  Icon:{
    Url: string;
    Description: string;
  };
  URL: any;
  BottomLinkIcon:{
    Url: string;
    Description: string;
  };

}

export default class ZentalisDepBannerWebPart extends BaseClientSideWebPart<IZentalisDepBannerWebPartProps> {

  private _getListData(): Promise<ISPLists> {
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.properties.ListName}')/Items?$select=Title,SubTitle,TilesBackgroundColor,TitleColor,SubTitleColor,Links,Icon,SectionPosition,IsActive,OrderID,BottomLinkIcon&$filter=IsActive eq 1`;
    return this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .catch(error => {
        console.error("Failed to fetch list data: ", error);
        return { value: [] }; // return empty list on error
      });
  }

  private _renderListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint ||
        Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getListData().then((response) => {
        this._renderList(response.value);
      });
    }
  }

  private _renderList(items: ISPList[]): void {
   
   
  
    let topHtml: string = '';
    let bottomHtml: string = '';
  
    items.forEach((item: ISPList) => {
      const iconUrl = item.Icon && item.Icon.Url ? escape(item.Icon.Url) : ''; // Check if item.Icon and item.Icon.Url are defined
  
      const linksContainerClass = item.SectionPosition === 'Bottom' ? 'Links_container_Bottom' : 'Links_container';
      const middleHeaderClass = item.SectionPosition === 'Bottom' ? 'Middle_header_Bottom' : 'Middle_header';
  
      // Check if BottomLinkIcon is defined before accessing its Url
      const bottomLinkIconUrl = item.BottomLinkIcon?.Url || '';
      const bottomLinkIconHtml = bottomLinkIconUrl ? `<img class="BottomLinkIcon" src="${escape(bottomLinkIconUrl)}" alt="Icon" />` : '';
  
      const itemHtml = `
      <div class="${linksContainerClass}" style="background-color:${escape(item.TilesBackgroundColor)};">
        <div class="${middleHeaderClass}">
          <div class="Hr">
            <a href="${escape(item.Links.Url)}" class="hover-container">
              <p style="color:${escape(item.SubTitleColor)};">${escape(item.SubTitle)}</p>
              <div class="flexContainer">
                ${iconUrl ? `<img class="topBannerLinkIcon" src="${iconUrl}" alt="Icon" />` : ''}
                <h2 class="${item.SectionPosition === 'Bottom' ? 'bottomSectionTitle' : 'topSectionTitle'}" style="color:${escape(item.TitleColor)};">
                  ${escape(item.Title)}
                </h2>
                ${bottomLinkIconHtml}
              </div>
            </a>
          </div>
        </div>
      </div>
    `;
    
    
  
      if (item.SectionPosition === 'Top') {
        topHtml += itemHtml;
      } else if (item.SectionPosition === 'Bottom') {
        bottomHtml += itemHtml;
      }
    });
  
    let html: string = `
      <div class="topBannerContainer">
        <div class="topBannerImage">
          <div class="topBannerText">${this.properties.WelcomeText}</div>
          <div class="topBannerDate"><p id="display-date"></p></div>
        </div>
      </div>
    `;
  
    html += `<div class="topContainer">${topHtml}</div>`;
    html += `<div class="bottomContainer">${bottomHtml}</div>`;
  
    const listContainer: Element = this.domElement.querySelector('#BindspListItems')!;
    listContainer.innerHTML = html;
  
    this._updateTiming();
    setInterval(() => this._updateTiming(), 60000);
  }
  
  private _updateTiming(): void {
    const today = new Date();
  
    // Adjusting for IST (UTC+5:30)
    const ISTOffset = 5.5 * 60; // in minutes
    const ISTTime = new Date(today.getTime() + ISTOffset * 60 * 1000); // Convert offset to milliseconds and add to current time
  
    const formattedDate = ISTTime.toDateString();
  
    const options: Intl.DateTimeFormatOptions = {
      hour: '2-digit',
      minute: '2-digit',
      hour12: true,
      timeZone: 'Asia/Kolkata' // Set the time zone to IST
    };
  
    let formattedTime = ISTTime.toLocaleTimeString('en-US', options);
    formattedTime = formattedTime.replace('AM', 'am').replace('PM', 'pm');
  
    const cleanedString = `${formattedDate} • ${formattedTime} IST`; // Append "IST" to the formatted time
    const displayDateElement = this.domElement.querySelector('#display-date');
  
    if (displayDateElement) {
      displayDateElement.innerHTML = cleanedString;
    }
  }

  public render(): void {
    const css = `
    
    `;
  
    // Add the CSS to the head of the document
    const style = document.createElement('style');
    style.type = 'text/css';
    style.appendChild(document.createTextNode(css));
    document.head.appendChild(style);
  
    this.domElement.innerHTML = `<div id="BindspListItems"></div>`;
    this._updateTiming(); // Call your updated time function here
    setInterval(() => this._updateTiming(), 60000); // Update time every minute
    this._renderListAsync();
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
                PropertyPaneTextField('WelcomeText', {
                  label: 'Welcome Text'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
