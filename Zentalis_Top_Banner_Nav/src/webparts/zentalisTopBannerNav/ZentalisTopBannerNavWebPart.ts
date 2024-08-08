import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'ZentalisTopBannerNavWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader'; 
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import * as moment from 'moment-timezone';

export interface IZentalisTopBannerNavWebPartProps {
  CSSUrl: string;
  WelcomeText: string;
  TopLink1: string;
  TopSubLink1: string;
  TopLinkUrl1: string;
  TopLink2: string;
  TopSubLink2: string;
  TopLinkUrl2: string;
  TopLink3: string;
  TopSubLink3: string;
  TopLinkUrl3: string;
  BottomLink1: string;
  BottomIcon1: string;
  BottomLinkUrl1: string;
  BottomLink2: string;
  BottomIcon2: string;
  BottomLinkUrl2: string;
  BottomLink3: string;
  BottomIcon3: string;
  BottomLinkUrl3: string;
  BottomLink4: string;
  BottomIcon4: string;
  BottomLinkUrl4: string;
  BottomLink5: string;
  BottomIcon5: string;
  BottomLinkUrl5: string;
  Icon: string;
}

export default class ZentalisTopBannerNavWebPart extends BaseClientSideWebPart<IZentalisTopBannerNavWebPartProps> {

  private _updateTiming(): void {
    const today = moment().tz(moment.tz.guess());
    const formattedDate = today.format('dddd, MMM DD YYYY · hh:mm a z');

    const displayDateElement = this.domElement.querySelector('#display-date');

    if (displayDateElement) {
      displayDateElement.innerHTML = formattedDate;
    }
  }
  
  public render(): void {
    const cssUrl = this.properties.CSSUrl || ''; 
    
    if (cssUrl) {
      SPComponentLoader.loadCss(cssUrl);
    }
    const userFullName = this.context.pageContext.user.displayName || "User";
    const userFirstName = escape(userFullName.split(' ')[0]);

    this.domElement.innerHTML = `
    <div class="navbar">
        <div class="header-section">
          <div class="container">
            <div class="header-text">
              <h1>${this.properties.WelcomeText}, <span>${userFirstName}!</span></h1>
               <p id="display-date"></p>
            </div>
          </div>
        </div>
    </div>
    <div class="Links_container">
        <div class="Middle_header">
          <div class="Hr">
            <p>${this.properties.TopSubLink1}</p>
            <a href="${this.properties.TopLinkUrl1}">
              <h2>${this.properties.TopLink1}</h2>
            </a>
          </div>
          <div class="Pto">
            <p>${this.properties.TopSubLink2}</p>
            <a href="${this.properties.TopLinkUrl2}">
              <h2>${this.properties.TopLink2}</h2>
            </a>
          </div>
          <div class="Staff">
            <p>${this.properties.TopSubLink3}</p>
            <a href="${this.properties.TopLinkUrl3}">
              <h2>${this.properties.TopLink3}</h2>
            </a>
          </div>
        </div>
        <div class="header_nav_links">
          <div class="logo_link head_col_3">
            <img src="${this.properties.BottomIcon1}" />
            <a href="${this.properties.BottomLinkUrl1}">
            ${this.properties.BottomLink1}<span class"NavigationIcon"><img src="${this.properties.Icon}" /></span>
            </a>
          </div>
          <div class="logo_link head_col_3">
            <img src="${this.properties.BottomIcon2}" />
            <a href="${this.properties.BottomLinkUrl2}">
            ${this.properties.BottomLink2}<span class"NavigationIcon"><img src="${this.properties.Icon}" /></span>
            </a>
          </div>
          <div class="logo_link head_col_3">
            <img src="${this.properties.BottomIcon3}" />
            <a href="${this.properties.BottomLinkUrl3}">
            ${this.properties.BottomLink3}<span class"NavigationIcon"><img src="${this.properties.Icon}" /></span>
            </a>
          </div>
          <div class="logo_link head_col_3">
            <img src="${this.properties.BottomIcon4}" />
            <a href="${this.properties.BottomLinkUrl4}">
            ${this.properties.BottomLink4}<span class"NavigationIcon"><img src="${this.properties.Icon}" /></span>
            </a>
          </div>
          <div class="logo_link head_col_3">
            <img src="${this.properties.BottomIcon5}" />
            <a href="${this.properties.BottomLinkUrl5}">
            ${this.properties.BottomLink5}<span class"NavigationIcon"><img src="${this.properties.Icon}" /></span>
            </a>
          </div>
        </div>
      </div>

    `;

    this._updateTiming();  
    setInterval(() => this._updateTiming(), 60000);  
    this._renderListAsync();
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
    });
  }

  private _renderListAsync(): void {
    if (Environment.type === EnvironmentType.SharePoint ||
        Environment.type === EnvironmentType.ClassicSharePoint) {
    }
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
          groups: [
            {
              groupName: 'Global CSS Path',
              groupFields: [
                PropertyPaneTextField('CSSUrl', {
                  label: 'CSS Url'
                })
              ]
            },
            {
              groupName: 'Top Navigation',
              groupFields: [
                PropertyPaneTextField('WelcomeText', {
                  label: 'Welcome Text'
                }),
              ]          
            },
            {
              groupName: 'Top Links',
              groupFields: [
                PropertyPaneTextField('TopLink1', {
                  label: 'Top Link Title 1'
                }),
                PropertyPaneTextField('TopLinkUrl1', {
                  label: 'Top Link Url 1'
                }),
                PropertyPaneTextField('TopSubLink1', {
                  label: 'Top Link Sub Title 1'
                }),                
                PropertyPaneTextField('TopLink2', {
                  label: 'Top Link Title 2'
                }),
                PropertyPaneTextField('TopLinkUrl2', {
                  label: 'Top Link Url 2'
                }),
                PropertyPaneTextField('TopSubLink2', {
                  label: 'Top Link Sub Title 2'
                }),                
                PropertyPaneTextField('TopLink3', {
                  label: 'Top Link Title 3'
                }),
                PropertyPaneTextField('TopLinkUrl3', {
                  label: 'Top Link Url 3'
                }),
                PropertyPaneTextField('TopSubLink3', {
                  label: 'Top Link Sub Title 3'
                })
                
              ]
            },
            {
              groupName: 'Bottom Links',
              groupFields: [
                PropertyPaneTextField('BottomLink1', {
                  label: 'Bottom Link Title 1'
                }),
                PropertyPaneTextField('BottomLinkUrl1', {
                  label: 'Bottom Link Url 1'
                }),
                PropertyPaneTextField('BottomIcon1', {
                  label: 'Bottom Link Icon 1'
                }),                
                PropertyPaneTextField('BottomLink2', {
                  label: 'Bottom Link Title 2'
                }),
                PropertyPaneTextField('BottomLinkUrl2', {
                  label: 'Bottom Link Url 2'
                }),
                PropertyPaneTextField('BottomIcon2', {
                  label: 'Bottom Link Icon 2'
                }),                
                PropertyPaneTextField('BottomLink3', {
                  label: 'Bottom Link Title 3'
                }),
                PropertyPaneTextField('BottomLinkUrl3', {
                  label: 'Bottom Link Url 3'
                }),
                PropertyPaneTextField('BottomIcon3', {
                  label: 'Bottom Link Icon 3'
                }),               
                PropertyPaneTextField('BottomLink4', {
                  label: 'Bottom Link Title 4'
                }),
                PropertyPaneTextField('BottomLinkUrl4', {
                  label: 'Bottom Link Url 4'
                }),
                PropertyPaneTextField('BottomIcon4', {
                  label: 'Bottom Link Icon 4'
                }),                
                PropertyPaneTextField('BottomLink5', {
                  label: 'Bottom Link Title 5'
                }),
                PropertyPaneTextField('BottomLinkUrl5', {
                  label: 'Bottom Link Url 5'
                }),
                PropertyPaneTextField('BottomIcon5', {
                  label: 'Bottom Link Icon 5'
                }),
                PropertyPaneTextField('Icon', {
                  label: 'Navigation Icon'
                })                
              ]
            }
          ]
        }
      ]
    };
  }
}
