import * as React from 'react';
import { TextField, DefaultButton, PrimaryButton, DialogFooter, Panel, Spinner, SpinnerType,DatePicker,Dropdown, DropdownMenuItemType, IDropdownOption } from "office-ui-fabric-react";

import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

import { ICustomPanelProps } from '../extensions/pp/PpCommandSet';

import { SPHttpClient, SPHttpClientResponse,ISPHttpClientOptions } from '@microsoft/sp-http';

export interface ICustomPanelState {
    saving: boolean;
   }


// Function to get the request digest value from SharePoint

/*
const getRequestDigest = async (httpClient: SPHttpClient, siteUrl: string): Promise<string> => {
    const requestOptions: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      }
    };
  
    const response: SPHttpClientResponse = await httpClient.post(`${siteUrl}/_api/contextinfo`, SPHttpClient.configurations.v1, requestOptions);
    const data = await response.json();
    return data.FormDigestValue;
  }; 
  */




const dllopt: IDropdownOption[] = [
    { key: 'StatusHeader', text: 'Status', itemType: DropdownMenuItemType.Header },
    { key: 'Open', text: 'Open' },
    { key: 'Pending', text: 'Pending',disabled: true },
    { key: 'In-Progress', text: 'In-Progress',disabled: true },
    { key: 'Closed', text: 'Closed',disabled: true },
    { key: 'divider_1', text: '-', itemType: DropdownMenuItemType.Divider },
  ];


export default class CustomPanel extends React.Component<ICustomPanelProps, ICustomPanelState> {

    //private editedTitle: string = null;

    constructor(props: ICustomPanelProps) {
        super(props);
        this.state = {
            saving: false
            
              
        };
      
this.handleSaveClick = this.handleSaveClick.bind(this);


    }


    /*  
    @autobind
    private _onTitleChanged(title: string) {
        this.editedTitle = title;
    }

    @autobind
    

    @autobind
    private _onSave() {
        this.setState({ saving: true });
        sp.web.lists.getById(this.props.listId).items.getById(this.props.itemId).update({
            'Title': this.editedTitle
        }).then(() => {
            this.setState({ saving: false });
            this.props.onClose();
        });
    }
 */
   


private handleSaveClick() {
   this._Save(); // Call fetchData method when Save button is clicked
   

}

     async _Save() {

        console.info(`http value a : ${this.props.spHttpClient}`);
        console.info(`${this.props.SiteURL}/_api/web/lists/getbytitle('Child Risk Status')/items`);
       
        const options: ISPHttpClientOptions = {
            headers: { 
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=verbose',
                'odata-version': ''},
            body: JSON.stringify({
                    '__metadata': {
                      'type': 'SP.Data.ActionItem_x005f_TaskListItem'
                    },
                    'Title':(document.getElementById('MainDesc') as HTMLInputElement).value,
                    'Main_Item': (document.getElementById('itemID') as HTMLInputElement).value
            })
            };
    
        try {
            const response: SPHttpClientResponse = await this.props.spHttpClient.post(
                `${this.props.SiteURL}/_api/web/lists/getbytitle('Child%20Risk%20Status')/items`,
                SPHttpClient.configurations.v1,
                options
            );
    
            if (response.ok) {
                const data = await response.json();
                return data.Id; // return the ID of the newly added item
            } else {
                console.error(`Failed to save item. Error: ${response.statusText}`);
                return -1; // return a default value indicating failure
            }
        } catch (error) {
            console.error(`An error occurred while saving item: ${error}`);
            return -1; // return a default value indicating failure
        }
    }


    private _onCancel() {
       window.parent.location.reload();

    }

    private _getPeoplePickerItems(items: any[]) {
        console.log('Items:', items);
      }

    public render(): React.ReactElement<ICustomPanelProps> {

        let { 
            isOpen,
            currentTitle,
            itemId
        } = this.props;


        return (
            <Panel isOpen={isOpen}>
                <h2>Child Item Form : {currentTitle}</h2>
                <TextField value={itemId.toString()} id='itemID' disabled={true}/>
                <TextField value={currentTitle} id='MainDesc'  label="Task" multiline={true} placeholder="Choose the new title" disabled={true} />
                <DatePicker placeholder="Select a date..." label="Date Time" ariaLabel="Select a date" />
                
                

                    
                    <Dropdown placeholder="Select Status" label="Status" options={dllopt}  />

                {this.state.saving && <Spinner type={SpinnerType.large} label="Saving…" />}
                <DialogFooter>
                    <DefaultButton text="Cancel" onClick={this._onCancel}  />
                    <PrimaryButton text="Save" onClick={this.handleSaveClick} />
                </DialogFooter>
            </Panel>
        );
    }
}


