import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as moment from 'moment-timezone';
//import styles from './ZentalisEventsWebPart.module.scss';
import * as strings from 'ZentalisEventsWebPartStrings';

// import { SPComponentLoader } from '@microsoft/sp-loader'; 
// const cssUrl = `https://realitycraftprivatelimited.sharepoint.com/sites/DevJay/SiteAssets/Zentalis.css`;
// SPComponentLoader.loadCss(cssUrl);

export interface IZentalisEventsWebPartProps {
  description: string;
  Calendar: string;
  UpcomingEvents: string;
  ViewAllEvents: string;
  ListName: string;
  RightArrow: string;
 
}

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  EventTitle: string;
  EventDate: string;
  Event_x0020_Category: string;
  EventLink: {
    Url: string;
    Description: string;
  };
}

export default class ZentalisEventsWebPart extends BaseClientSideWebPart<IZentalisEventsWebPartProps> {

  private _getListData(): Promise<ISPLists> {
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.properties.ListName}')/items?$select=EventTitle,EventDate,Event_x0020_Category,EventLink`;
    console.log('Request URL:', requestUrl);

    return this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        console.log('Response status:', response.status, response.statusText);
        if (!response.ok) {
          throw new Error(`Error fetching data. Status: ${response.statusText}`);
        }
        return response.json();
      })
      .catch((error) => {
        console.error('Error fetching data:', error);
        throw error;
      });
  }

  private _formatEventDate(eventDate: string): string {
    const date = moment(eventDate).tz(moment.tz.guess());
    const daysOfWeek = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const monthsOfYear = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  
    const dayOfWeek = daysOfWeek[date.day()];
    const month = monthsOfYear[date.month()];
    const dayOfMonth = date.date();
    let hours = date.hours();
    const minutes = date.minutes();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12; // Handle midnight
    const minutesStr = minutes < 10 ? '0' + minutes : minutes.toString();
    const timezone = date.format('z'); // Get the timezone abbreviation
  
    return `${dayOfWeek}, ${month} ${dayOfMonth}, ${hours}:${minutesStr} ${ampm} (${timezone})`;
  }

  private _renderListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint ||
        Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          if (response && response.value) {
            console.log('Response data:', response.value);
            this._renderList(response.value);
          } else {
            console.warn('No data found in response.');
            this._renderList([]);
          }
        })
        .catch((error) => {
          console.error('Error in _renderListAsync:', error);
          this._renderList([]);
        });
    }
  }

  private _renderList(items: ISPList[]): void {
    items.sort((a, b) => new Date(a.EventDate).getTime() - new Date(b.EventDate).getTime());
    const today = new Date();
    items = items.filter(item => new Date(item.EventDate) >= today);
    items = items.slice(0, 6);

    let html: string = `
      <div class="calendarContainer">
        <div class="calendarText">
          <div class="calendarLeft">
            <h5>${this.properties.Calendar}</h5>
            <h1>${this.properties.UpcomingEvents}</h1>
          </div>
          <div class="calendarRight">
            <a href="#">${this.properties.ViewAllEvents}</a>
            <img src="${this.properties.RightArrow}" alt="RightArrow"/>
          </div>
        </div>
        <div class="calendarAllEvents">
    `;

    if (items.length > 0) {
      items.forEach((item: ISPList, index: number) => {
        const eventDate = new Date(item.EventDate);
        const month = eventDate.toLocaleString('default', { month: 'short' }).toUpperCase();
        const day = eventDate.getDate();

        if (index % 3 === 0) {
          if (index > 0) {
            html += `</div>`;
          }
          html += `<div class="calendarRow">`;
        }

        html += `
          <div class="calendarEvents">
            <div class="calendarDates">
              <p>${month}</p>
              <p class="calendarNumber">${day}</p>
            </div>
            <div class="calendarEventsItem">
              <a href="${item.EventLink.Url}"><p class="depart">${item.Event_x0020_Category}</p></a>
              <p class="calendarEventName">${item.EventTitle}</p>
              <p class="datetime">${this._formatEventDate(item.EventDate)}</p>
            </div>
          </div>
        `;
      });

      if (items.length % 3 !== 0) {
        html += `</div>`;
      }
    } else {
      html += `<p>No upcoming events</p>`;
    }

    html += `</div></div>`;

    const listContainer: Element = this.domElement.querySelector('#BindspListItems')!;
    listContainer.innerHTML = html;
  }

  public render(): void {

    this.domElement.innerHTML = `<div class="calendarTopContainer" id="BindspListItems"></div>`;
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
                PropertyPaneTextField('Calendar', {
                  label: 'Title'
                }),
                PropertyPaneTextField('UpcomingEvents', {
                  label: 'Upcoming Events'
                }),
                PropertyPaneTextField('ViewAllEvents', {
                  label: 'View All Events'
                }),
                PropertyPaneTextField('RightArrow', {
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
