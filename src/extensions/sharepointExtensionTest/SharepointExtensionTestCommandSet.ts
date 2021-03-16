import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  RowAccessor
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import customProtocolCheck from "custom-protocol-check";

import * as strings from 'SharepointExtensionTestCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISharepointExtensionTestCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'SharepointExtensionTestCommandSet';

export default class SharepointExtensionTestCommandSet extends BaseListViewCommandSet<ISharepointExtensionTestCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized SharepointExtensionTestCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_OPEN_NITRO');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length >= 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_OPEN_NITRO':
        event.selectedRows.forEach( element => {
          this.openPdfInNitro(this.getUrlForItem(element));
        });
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private getUrlForItem(item : RowAccessor) : string {
        //e.g. https://mysite.sharepoint.com/sites/Test
        const siteAbsUrl = this.context.pageContext.site.absoluteUrl;

        //e.g. /sites/Test
        const siteRelUrl = this.context.pageContext.site.serverRelativeUrl;
    
        //sites/Test/Shared Documents/MyDocument.pdf
        const docRelativeUrl :string = item.getValueByName('FileRef');
    
        //form the absolute url
        const absUrl = `${siteAbsUrl}` +`${docRelativeUrl.substr(siteRelUrl.length)}`;
    
        return absUrl;
  }

  private openPdfInNitro(item : string) : void {
    if (window.navigator.platform != "Win32") {
      Dialog.alert("Nitro extension is only supported on Windows");
      return;
    }

    var data = {
        "sharepoint" : {
          url: item,
        }
    };
    var customUrl = "sharepoint:" + JSON.stringify(data);
    customProtocolCheck( 
      customUrl,
      () => {
        Dialog.alert("Custom protocol not found");
      },
      () => {
        window.open(customUrl, "_blank");
      }
    );
  }
}
