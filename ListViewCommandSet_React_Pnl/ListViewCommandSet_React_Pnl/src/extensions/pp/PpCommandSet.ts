import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  IListViewCommandSetListViewUpdatedParameters,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import CustomPanel from '../../components/CustomPanel';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { assign } from 'office-ui-fabric-react';
import { SPHttpClient } from '@microsoft/sp-http';
import { BaseComponentContext } from '@microsoft/sp-component-base';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPpCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}


export interface ICustomPanelProps {
  onClose: () => void;
  isOpen: boolean;
  currentTitle: string;
  itemId: number;
  listId: string;
  SiteURL:string;
  spHttpClient: SPHttpClient;
  bcontext:BaseComponentContext;
  
}



const LOG_SOURCE: string = 'PpCommandSet';

export default class PpCommandSet extends BaseListViewCommandSet<IPpCommandSetProperties> {





  //Declare
  public panelPlaceHolder:HTMLDivElement ;
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized PpCommandSet');




    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    compareOneCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

     // Create the container for our React component
    this.panelPlaceHolder = document.body.appendChild(document.createElement("div"));
 
 return Promise.resolve();

  }

  private _showPanel(itemId: number, currentTitle: string) {
    this._renderPanelComponent({
      isOpen: true,
      currentTitle,
      itemId,
      listId: '',
      onClose: this._dismissPanel,
      SiteURL: this.context.pageContext.web.absoluteUrl,
      spHttpClient:this.context.spHttpClient,
      bcontext:this.context,
    });
  
  }
  

  private _dismissPanel() {
    this._renderPanelComponent({ isOpen: false });
  }

  private _renderPanelComponent(props: any) {
    const element: React.ReactElement<ICustomPanelProps> = React.createElement(CustomPanel, assign({
      onClose: null,
      currentTitle: null,
      itemId: null,
      isOpen: false,
      listId: null,
      SiteURL:null,
      spHttpClient:null,
      bcontext:null
    }, props));
    
    ReactDOM.render(element, this.panelPlaceHolder);
  }

  
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const openEditorCommand: Command = this.tryGetCommand('COMMAND_1');
    openEditorCommand.visible = event.selectedRows.length === 1;
  }



  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        let selectedItem = event.selectedRows[0];
        const listItemId = selectedItem.getValueByName('ID') as number;
        const title = selectedItem.getValueByName("Risk Statement") as string;
        const Status = selectedItem.getValueByName("Status") as string;
        console.log(`Item ID : ${listItemId} , Item Title : ${title} , Status: ${Status}`);
        this._showPanel(listItemId,title);
        //Dialog.alert(`${this.properties.sampleTextOne}`).catch(() => {  });



        break;
      case 'COMMAND_2':
        Dialog.alert(`${this.properties.sampleTextTwo}`).catch(() => {
          /* handle error */
        });
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = this.context.listView.selectedRows?.length === 1;
    }

    // TODO: Add your logic here

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();

  }




  
}