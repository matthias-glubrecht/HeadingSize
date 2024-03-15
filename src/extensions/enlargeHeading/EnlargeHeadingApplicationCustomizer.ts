import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'EnlargeHeadingApplicationCustomizerStrings';

const LOG_SOURCE: string = 'EnlargeHeadingApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IEnlargeHeadingApplicationCustomizerProperties {
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class EnlargeHeadingApplicationCustomizer
  extends BaseApplicationCustomizer<IEnlargeHeadingApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    this.addStyleToPage();
    return Promise.resolve();
  }

  private addStyleToPage() : void {
    const head: HTMLHeadElement = document.getElementsByTagName('head')[0];
    const style: HTMLStyleElement = document.createElement('style');
    //style.type = 'text/css';
    style.innerHTML = `
      #CaptionElementView {
        font-size: 24px !important;
      }
    `;
    head.appendChild(style);
  }
}
