import { Version } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import pnp, { setup as pnpSetup } from '@pnp/pnpjs';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SupplierContactsWebPart.module.scss';
import * as strings from 'SupplierContactsWebPartStrings';
import * as $ from "jquery";
const t = $;
import 'datatables.net';
import 'datatables.net-dt';
SPComponentLoader.loadCss("https://cdn.datatables.net/1.12.1/css/jquery.dataTables.min.css");
var cred = {username:"",password:""},renderThisCtrl: any;
export interface ISupplierContactsWebPartProps {
  description: string;
}


export default class SupplierContactsWebPart extends BaseClientSideWebPart<ISupplierContactsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  public table1:any;
  public render(): void {
    renderThisCtrl = this;
    this.domElement.innerHTML = `<table id="firstTable" style="display:none;">
                                      <thead style="display:none;">
                                        <tr role="row">
                                          <th></th>
                                          <th></th>
                                        </tr>
                                      </thead>
                                      <tbody>
                                        <tr>
                                          <td id="supHead"></td>
                                          <td id="supName"></td>
                                        </tr>
                                        <tr>
                                          <td></td>
                                          <td></td>
                                        </tr>
                                        <tr>
                                          <td id="genHead"></td>
                                          <td id="genEmail"></td>
                                        </tr>                        
                                      </tbody>
                                  </table>
                                  <table id="myTable" class="display table-responsive-md no-footer dataTable" role="grid">
                                    <thead>
                                      <tr role="row" id="theader">
                                        
                                      </tr>
                                    </thead>
                                    <tbody id="tblbodyAll">                        
                                    </tbody>
                                </table>`;
        var currentUser = this.context.pageContext.legacyPageContext.userEmail;
        var webUrl = this.context.pageContext.web.absoluteUrl;
        var siteName = this.context.pageContext.web.title;
        var siteId = this.context.pageContext.site.id.toString();
        this.getData(currentUser,webUrl,siteName,siteId);
  }

  private getData(user:string,webUrl:string,siteName:string,siteId:string){
   // var bodyContent = "{\r\n \"Email\":\"" + user + "\,\r\n \"SiteId\": \"\"\r\n}"; 
    var bodyContent = "{\r\n \"Email\":\"" + user + "\",\r\n \"SiteId\": \"" + siteId + "\",\r\n \"WebUrl\": \"" + webUrl + "\",\r\n \"SiteName\": \"" + siteName + "\"\r\n}";
    $.ajax({
      url: "https://prod-199.westeurope.logic.azure.com:443/workflows/b9ef21cf3719451480c5199cdcbdd9d2/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=i704iwXzTiLTtT85ijAxiIPc36pHVKL27Lc_2Sxaqxs",
      crossDomain: true,
      method: "POST",
      data: bodyContent,
      processData: false,
      headers: {
          "content-type": "application/json",
          "cache-control": "no-cache"
      },
      success: function (data) {
          if(data){
            var pass = btoa(data.Key + ':' + data.Value),isheader = false;
            if(data.MainHeader && data.MainHeader.length>0){
              $('#firstTable').show();
              $('#supHead').html('<b>' + data.MainHeader[0].Title || "" + ': </b>');
              $('#genHead').html('<b>' + data.MainHeader[1].Title || "" + ': </b>');
              isheader = true;
            }
            var headerColumnsArray = data.DataRows;
            $.ajax
            ({
              type: "GET",
              url: "https://sapdev-mercedesamgf1.msappproxy.net/sap/opu/odata/sap/ZODATA_SUPP_PORTAL_SRV/ContactDetailsSet?$filter=(SupplierCode eq '" + data.SupplierCode + "')&$expand=ContactAlternatives&sap-client=700",
              headers: {
                "Authorization": "Basic " + pass,
                "Content-Type": "application/json",
                "Accept": "application/json"
              },        
              crossDomain: true,
              success: function (data){
                console.log(data);
                if(data.d.results.length > 0 && data.d.results[0].ContactAlternatives.results.length > 0){
                  if (renderThisCtrl.table1 instanceof (<any>$.fn.dataTable).Api) {
                    $('#myTable').DataTable().clear();
                    $('#myTable').DataTable().destroy();
                  } 
                  if(isheader){
                    $('#supName').html(data.d.results[0].SupplierName || "");
                    $('#genEmail').html(data.d.results[0].GeneralContactEmail || "");
                  }
                  var headerArray = Object.keys(data.d.results[0].ContactAlternatives.results[0]).filter(val => val != "__metadata" && val != "AltContactTelCode" && val != "SupplierCode");
                  $('#theader').html('');
                  $('#tblBodyAll').html('');
                  for (let index = 0; index < headerArray.length; index++) {
                    const element = headerArray[index];
                    if(index == 0)
                      $('#theader').append('<th class="sorting_asc" tabindex="0" aria-controls="tblAll" rowspan="1" colspan="1" aria-label="File: activate to sort column descending" aria-sort="ascending">' + headerColumnsArray.filter(function (el:any){return el.Name ==element;})[0].Title || "" + '</th>');
                    else
                    $('#theader').append('<th class="sorting" tabindex="0" aria-controls="tblAll" rowspan="1" colspan="1" aria-label="File: activate to sort column descending" aria-sort="ascending">' + headerColumnsArray.filter(function (el:any){return el.Name ==element;})[0].Title || "" + '</th>');
                  }
                  for (let index2 = 0; index2 < data.d.results[0].ContactAlternatives.results.length; index2++) {
                    var element2 = data.d.results[0].ContactAlternatives.results[index2];//<td>` + element.SupplierCode + `</td>
                    var tr = `<tr role="row" class="odd">`;
                    for (let index = 0; index < headerArray.length; index++) {
                      if(index == 0)
                        tr += `<td class="sorting_1">` + element2[headerArray[index]] || "" + `</td>`;
                      else
                        tr += `<td>` + element2[headerArray[index]] || "" + `</td>`;                      
                    }
                    
                    tr += `</tr>`;
                    $('#tblbodyAll').append(tr);
                  }
                  renderThisCtrl.table1 = $('#myTable').DataTable({
                    paging: true,
                    ordering: true,
                    processing: true,
                    "lengthChange": false,
                    "info": false
                  });
                }
                //$('#sapData').html(data.d.results[0].SAPResponse);
              },
              error:function(err){
                console.log(err);
              }
            });
          }
      },
      error: function (err) {
        console.log(err);
      }
    });           
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }



  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
