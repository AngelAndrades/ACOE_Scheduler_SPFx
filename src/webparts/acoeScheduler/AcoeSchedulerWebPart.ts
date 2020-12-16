import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

//import styles from './AcoeSchedulerWebPart.module.scss';
//import * as strings from 'AcoeSchedulerWebPartStrings';

import * as $ from 'jquery';
import '@progress/kendo-ui';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as pnp from '@pnp/sp/presets/all';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { SPA } from './apps/spa';

export interface IAcoeSchedulerWebPartProps {
  collectionData: any[];
}

export default class AcoeSchedulerWebPart extends BaseClientSideWebPart<IAcoeSchedulerWebPartProps> {

  public render(): void {
    /*
        Load external CDN files (JS/CSS)
    */
    SPComponentLoader.loadCss('https://kendo.cdn.telerik.com/2020.3.915/styles/kendo.common-material.min.css');
    SPComponentLoader.loadCss('https://kendo.cdn.telerik.com/2020.3.915/styles/kendo.material.min.css');

    SPComponentLoader.loadScript('https://kendo.cdn.telerik.com/2020.3.915/js/jszip.min.js');

    if (this.properties.collectionData == null) this.domElement.innerHTML = `<h3>Edit Web Part Property Pane</h3><p>Go to the web part property page and add new calendars to the Schedule Viewer. You are going to need the following information for each calendar:</p><ul><li>Calendar Name:  this is the name you see on the web page</li><li>Site URL:  enter only the site url, do not include library names</li><li>Color (hex values only): Use this <a href="https://color.adobe.com/create/color-wheel" title="link to Adobe Color site" target="_blank">Adobe Color</a> site to select a complementary color palette</li></ul>`;
    else this.domElement.innerHTML = `<div><table id="filter" style="min-width: 100%; padding-bottom: 15px;"></table></div><div id="calendar"></div>`;

    const spa = SPA.getInstance(this.properties.collectionData);
  }

  protected onInit(): Promise < void > {
    return super.onInit().then(_ => {
      pnp.sp.setup({
        spfxContext: this.context,
        sp: {
          headers: {
            Accept: 'application/json;odata=nometadata'
          }
        }
      });
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Select the button below and enter the required data for each calendar'
          },
          groups: [
            {
              groupFields: [
                PropertyFieldCollectionData('collectionData', {
                  key: 'collectionData',
                  label: '',
                  panelHeader: 'Enter information for each calendars you want to combine',
                  manageBtnLabel: 'Add Calendars Information',
                  value: this.properties.collectionData,
                  fields: [
                    { id: 'Title', title: 'Calendar Name', type: CustomCollectionFieldType.string, required: true },
                    { id: 'URL', title: 'Site URL', type: CustomCollectionFieldType.string, required: true },
                    { id: 'Color', title: 'Color (hex value)', type: CustomCollectionFieldType.string, required: true }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
