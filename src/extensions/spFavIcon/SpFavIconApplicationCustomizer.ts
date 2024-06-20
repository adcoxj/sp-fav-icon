import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as _strings from 'SpFavIconApplicationCustomizerStrings';

const LOG_SOURCE: string = 'SpFavIconApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpFavIconApplicationCustomizerProperties {
  // This is an example; replace with your own property
  favicon: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpFavIconApplicationCustomizer
  extends BaseApplicationCustomizer<ISpFavIconApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    let url: string = this.properties.favicon;
    if (!url) {
      Log.info(LOG_SOURCE, `Fav Icon URL is missing.`);
    }else{
      let url = "https://centralhealthtx.sharepoint.com/sites/ENT-BrandCentral/Branding%20files/favicon.ico";
      const link = document.getElementById('favicon');
      if(link ==null)
      {
        const link = document.createElement("link");
        link.setAttribute('type', 'image/x-icon');
        link.setAttribute('rel', 'shortcut icon');
        link.setAttribute('href', url);
        let head = document.head;
        head.appendChild(link);
      }
      else{
        link.setAttribute('type', 'image/x-icon');
        link.setAttribute('rel', 'shortcut icon');
        link.setAttribute('href', url);
        let head = document.head;
        head.appendChild(link);
      }
    }

    return Promise.resolve();
  }
}
