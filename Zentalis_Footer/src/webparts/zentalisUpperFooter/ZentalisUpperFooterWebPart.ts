import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

//import styles from './ZentalisUpperFooterWebPart.module.scss';
import * as strings from 'ZentalisUpperFooterWebPartStrings';

// import { SPComponentLoader } from '@microsoft/sp-loader'; 
// const cssUrl = `https://realitycraftprivatelimited.sharepoint.com/sites/DevJay/SiteAssets/Zentalis.css`;
// SPComponentLoader.loadCss(cssUrl);

export interface IZentalisUpperFooterWebPartProps {
  description: string;
  ZentalisContact: string;
  ZentalisFottertitle: string;
  ZentalisFotterDescription: string;
  ZentalisSchedulingMeeting: string;
  FooterLogo: string;
  LowerFooterText: string;
  CopyrightText: string;
  RightArrow: string;
  ZentalisSchedulingMeetingLink: string;

}

export default class ZentalisUpperFooterWebPart extends BaseClientSideWebPart<IZentalisUpperFooterWebPartProps> {


  public render(): void {
    this.domElement.innerHTML = `
    <div class="upperFooterContainer">
      <div class="footer_container">
        <div class="uppercontainer">
          <div class="footer">
            <div class="div1">
              <div class="div1_up">
                <div class="div1_up_upper">
                  <div class="F1"></div>
                  <div class="F2"></div>
                </div>
                <div class="div1_up_lower">
                  <div class="F3"></div>
                </div>
              </div>
              <div class="div1_low">
                <div class="div1_low_upper">
                  <div class="F4"></div>
                  <div class="F5"></div>
                </div>
                <div class="div1_low_lower">
                  <div class="F6"></div>
                </div>
              </div>
            </div>
            <div class="div2">
              <div class="F7">
                <div class="footer_two_middle">
                  <p class="title_one">${this.properties.ZentalisContact}</p>
                  <h2 class="title_two">${this.properties.ZentalisFottertitle}</h2>
                  <p class="title_three">${this.properties.ZentalisFotterDescription}</p>
                  
                </div>
              </div>
            </div>
            <div class="div3">
              <div class="div3_up">
                <div class="div3_up_upper">
                  <div class="F8"></div>
                </div>
                <div class="div3_up_lower" style="display: flex; justify-content: flex-end">
                  <div class="F9"></div>
                  <div class="F10"></div>
                </div>
              </div>
              <div class="footer_link_Bottom">
                    <a href="${this.properties.ZentalisSchedulingMeetingLink}">${this.properties.ZentalisSchedulingMeeting}</a>
                    <img src="${this.properties.RightArrow}" alt="" />
                  </div>
              <div class="div3_low">
                <div class="div3_low_upper" style="display: flex; justify-content: flex-end">
                  <div class="F11"></div>
                </div>
                <div class="div3_low_lower">
                  <div class="F12"></div>
                  <div class="F13"></div>
                </div>
              </div>
            </div>

            <footer>
            <div class="footeraddress">
              <div class="footerimg">
                <img src="${this.properties.FooterLogo}" alt="err-logo" />
              </div>
              <div class="footertexts">
                <p>${this.properties.CopyrightText}</p>
                <p>${this.properties.LowerFooterText}</p>
              </div>
            </div>
          </footer>
          </div>
        </div>
      </div>
    </div> `;
}

  

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
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

    const {
      semanticColors
    } = currentTheme;

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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('ZentalisContact', {
                  label: 'Zentalis Contact'
                }),
                PropertyPaneTextField('ZentalisFottertitle', {
                  label: 'Fotter Title'
                }),
                PropertyPaneTextField('ZentalisFotterDescription', {
                  label: 'Fotter Description'
                }),
                PropertyPaneTextField('ZentalisSchedulingMeeting', {
                  label: 'Zentalis Scheduling Meeting'
                }),
                PropertyPaneTextField('ZentalisSchedulingMeetingLink', {
                  label: 'Scheduling Meeting Link'
                }),
                PropertyPaneTextField('RightArrow', {
                  label: 'Right Arrow'
                }),
                PropertyPaneTextField('FooterLogo', {
                  label: 'Footer Logo'
                }),
                PropertyPaneTextField('CopyrightText', {
                  label: 'Copyright Text'
                }),
                PropertyPaneTextField('LowerFooterText', {
                  label: 'Lower Footer Text'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
