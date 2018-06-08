import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CatDsgSpWp1004OurTeamWebPart.module.scss';
import * as strings from 'CatDsgSpWp1004OurTeamWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ICatDsgSpWp1004OurTeamWebPartProps {
  description: string;
  customFilter: string;
  membersLimit: number;
}
export interface ITeamMember {
  Title: string;
  JobTitle: string;
  PictureURL: string;
  WorkPhone: string;
  WorkEmail: string;
  Department: string;
  Ext: string;
  ProfileUrl: string;
}
export default class CatDsgSpWp1004OurTeamWebPart extends BaseClientSideWebPart<ICatDsgSpWp1004OurTeamWebPartProps> {

  public render(): void {
    this.properties.description = strings.CatDsgSpWp1004OurTeamDescription;
    this.domElement.innerHTML = `
      <div class="${ styles.catDsgSpWp1004OurTeam}">
        <div class="${ styles.container}">
          <ul class="${styles.catDsgSpWp1004OurTeamMembersContainer}">
          </ul>
        </div>
      </div>`;
    this.renderOurTeam();
  }

  protected renderOurTeam() {
    this.getOurTeamMembers().then((TeamMembers: ITeamMember[]) => {
      let sortedteamMembers = TeamMembers.sort((t1, t2) => t1.Title.localeCompare(t2.Title));
      let ourTeamMembersHtml = '';
      let encodedId = this.context.instanceId + "_picture3Lines_";
      var pictureId = encodedId + "picture";
      var containerId = encodedId + "container";
      var pictureLinkId = encodedId + "pictureLink";
      var pictureContainerId = encodedId + "pictureContainer";
      var dataContainerId = encodedId + "dataContainer";
      var line1LinkId = encodedId + "line1Link";
      var line1Id = encodedId + "line1";
      var line2Id = encodedId + "line2";
      var line3Id = encodedId + "line3";
      var line4Id = encodedId + "line4";
      if (sortedteamMembers.length > 0) {
        for (var j = 0; j < sortedteamMembers.length; j++) {
          let ourTeamMemberHtml = ` 
          <li>
            <div class="${styles["catDsgSpWp1004OurTeamPicture3LinesContainer"]}" id="${containerId}">
              <div class="${styles["catDsgSpWp1004OurTeamDtContainer"]}" id="${pictureContainerId}">` + (sortedteamMembers[j].ProfileUrl ?
              `<a class="${styles["catDsgSpWp1004OurTeamPictureImgLink"]}" href="${sortedteamMembers[j].ProfileUrl}" title="${sortedteamMembers[j].Title}" id="${pictureLinkId}">
                      <img class="${styles["catDsgSpWp1004OurTeamDtPicture"]}" src="${sortedteamMembers[j].PictureURL}"> 
                  </a>`: ``) + `
              </div>
              <div class="${styles["catDsgSpWp1004OurTeamPicture3LinesDataContainer"]}" id="${dataContainerId}">
                  <a class="${styles["catDsgSpWp1004OurTeamPicture3LinesLine1Link"]}" href="${sortedteamMembers[j].ProfileUrl}" title="${sortedteamMembers[j].Title}" id="${line1LinkId}">
                      <h2 class="${styles["catDsgSpWp1004OurTeamDtHeader"]}" id="${line1Id}"> ${sortedteamMembers[j].Title}</h2>
                  </a>`+ (sortedteamMembers[j].JobTitle ?
              `<div class="${styles["catDsgSpWp1004OurTeamDtLineItem"]}" title="${sortedteamMembers[j].JobTitle}" id="${line2Id}"> ${sortedteamMembers[j].JobTitle}</div>` :
              `<div class="${styles["catDsgSpWp1004OurTeamDtLineItem"]}" id="${line2Id}"> </div>
                  `) + (sortedteamMembers[j].WorkPhone ?
              `<div class="${styles["catDsgSpWp1004OurTeamDtLineItem"]}" id="${line3Id}" title="${sortedteamMembers[j].WorkPhone}">Tel. ${sortedteamMembers[j].WorkPhone}</div>` :
              `<div class="${styles["catDsgSpWp1004OurTeamDtLineItem"]}" id="${line3Id}">Tel. </div>
                  `) + (sortedteamMembers[j].Ext ?
              `<div class="${styles['catDsgSpWp1004OurTeamDtLineItem']}" id="${line4Id}" title="ext">Ext. ${sortedteamMembers[j].Ext}</div>` : '') + `
              </div>
            </div>
          </li>
          `;
          ourTeamMembersHtml += ourTeamMemberHtml;
        }
        this.domElement.querySelector("." + styles.catDsgSpWp1004OurTeamMembersContainer).innerHTML = ourTeamMembersHtml;
      }
      else {
        if (TeamMembers.length == 0) {
          this.context.statusRenderer.renderError(this.domElement, strings.CatDsgSpWp1004OurTeamNoDataMesssage);
        }
      }
    }, (error: any) => {
      this.context.statusRenderer.renderError(this.domElement, error);
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
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
                PropertyPaneTextField('customFilter', {
                  label: strings.CatDsgSpWp1004OurTeamFieldLabelCustomFilter
                }),
                PropertyPaneTextField('membersLimit', {
                  label: strings.CatDsgSpWp1004OurTeamFieldLabelMembersLimit,
                  onGetErrorMessage: this.validateMembersLimitProperty.bind(this)
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private validateMembersLimitProperty(value: string): string {
    if(value==null || (value!=null&&value.trim()=="") || isNaN(value as any)) {
      return strings.CatDsgSpWp1004OurTeamPropertyValidationErrorMessageValueMustBeGreaterThanZero;
    }
    let number: any = parseInt(value, 10);
    if (number <= 0) {
      return strings.CatDsgSpWp1004OurTeamPropertyValidationErrorMessageValueMustBeGreaterThanZero;
    }
    return "";

  }

  private getOurTeamMembers(): Promise<ITeamMember[]> {
    return new Promise<ITeamMember[]>((resolve, reject) => {
      this.getOurTeamMembersSearchingResult().then((searchResult: any) => {
        if (searchResult != undefined && searchResult != null) {
          var results = searchResult.PrimaryQueryResult.RelevantResults.Table.Rows;
          let TeamMembers: ITeamMember[] = [];
          for (var i = 0; i < results.length; i++) {
            let workEmail = results[i].Cells[9].Value;
            let pictureURL = results[i].Cells[3].Value;
            if (pictureURL) {
              pictureURL = this.context.pageContext.web.absoluteUrl + "/_layouts/15/userphoto.aspx?size=M&url=MThumb.jpg&accountname=" + workEmail;
            }
            else {
              pictureURL = this.context.pageContext.web.absoluteUrl + "/_layouts/15/userphoto.aspx?size=M&url=MThumb.jpg";
            }
            let teamMember: ITeamMember = {
              PictureURL: pictureURL,
              ProfileUrl: results[i].Cells[5].Value as string,
              Title: results[i].Cells[6].Value as string,
              JobTitle: results[i].Cells[7].Value as string,
              WorkPhone: results[i].Cells[8].Value as string,
              WorkEmail: results[i].Cells[9].Value as string,
              Department: results[i].Cells[10].Value as string,
              Ext: results[i].Cells[12].Value as string
            };
            TeamMembers.push(teamMember);
          }
          resolve(TeamMembers);
        }
      }, (error: any) => {
        return reject(error);
      });
    });
  }
  private getOurTeamMembersSearchingResult(): Promise<any> {
    const queryText: string = `querytext='${this.properties.customFilter ? this.properties.customFilter : '*'}'`;
    const selectedProperties: string = `selectproperties='PublishingImage,PictureURL,PictureThumbnailURL,Path,Title,JobTitle,WorkPhone,WorkEmail,Department,Extension'`;
    let url = `${this.context.pageContext.web.absoluteUrl}/_api/search/query?${queryText}&${selectedProperties}&sourceid='B09A7990-05EA-4AF9-81EF-EDFAB16C4E31'&rowlimit=${this.properties.membersLimit.toString()}`;
    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 404) {
          this.context.statusRenderer.renderError(this.domElement, strings.CatDsgSpWp1004OurTeam404Messsage);
          return [];
        }
        else if (response.status === 400) {
          this.context.statusRenderer.renderError(this.domElement, strings.CatDsgSpWp1004OurTeamBadRequestMessagePrefix + url);
          return [];
        }
        else {
          return response.json();
        }
      }, (error: any) => {
        this.context.statusRenderer.renderError(this.domElement, error);
      });
  }
}
