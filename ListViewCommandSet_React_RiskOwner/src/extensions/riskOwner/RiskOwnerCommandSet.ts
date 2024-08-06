import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */


export interface IRiskOwnerCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
  sampleTextthree: string;
  sampleTextfour: string;
  URL:string;
  State:String;
  StateOwner:String;
  
  Token:string;
  StateEMC:String;
  StateRMOC:String;

  Status:String;

}

const LOG_SOURCE: string = 'RiskOwnerCommandSet';

export default class RiskOwnerCommandSet extends BaseListViewCommandSet<IRiskOwnerCommandSetProperties> {

  public onInit(): Promise<void> {
  
  console.info('Start:');
  console.info(`Current User Email: ${this.context.pageContext.user.email}`);
    //Hide on onInit
    //#region Check Risk Owner
          const comparethreeCommand: Command = this.tryGetCommand('COMMAND_3');
          comparethreeCommand.visible = false;
          const comparefourCommand: Command = this.tryGetCommand('COMMAND_4');
          comparefourCommand.visible = false;
          this.GetRiskOwnerFromDivisionDepartmentBreakdown();
  //#endregion

    //#region Check Risk Officer
          const comparet5Command: Command = this.tryGetCommand('COMMAND_5');
          comparet5Command.visible = false;
          const compare6Command: Command = this.tryGetCommand('COMMAND_6');
          compare6Command.visible = false;
          this.GetRiskOfficerFromDivisionWiseApproval();
  //#endregion

    //#region Check RMOC Users
 const comparet7Command: Command = this.tryGetCommand('COMMAND_7');
          comparet7Command.visible = false;
          const compare8Command: Command = this.tryGetCommand('COMMAND_8');
          compare8Command.visible = false;
          this.GetMultipleRMOCFromDivisionWiseApproval();
 //#endregion


   //#region Check EMC Users
   const comparet9Command: Command = this.tryGetCommand('COMMAND_9');
   comparet9Command.visible = false;
   const compare10Command: Command = this.tryGetCommand('COMMAND_10');
   compare10Command.visible = false;
   this.GetMultipleEMCFromDivisionWiseApproval();
 //#endregion

Log.info(LOG_SOURCE, 'Initialized RiskOwnerCommandSet');

this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

return Promise.resolve();

}
  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {

    let selectedItem = event.selectedRows[0];
    const listItemId = selectedItem.getValueByName('ID') as string;
    
    let PAURLRiskOwner= 'https://prod-46.southeastasia.logic.azure.com:443/workflows/5840ebbe6bbe48508e67d9350433b6b2/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Wu6cY-GiZmJlfrC0zbcd9HHuhQKlE9-JQL7_9468VFk';
    let PAURLRiskOfficer= 'https://prod-17.southeastasia.logic.azure.com:443/workflows/fec9cc2b100e4bd386dd783807b6f410/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=tmsc60oY5cOaclG132zbGS6rw7iiHCb8TDGKwZ2_11M';
    let PAURLRiskRMOC= 'https://prod-05.southeastasia.logic.azure.com:443/workflows/b18299b7cecf4e0b8836b19e12476f21/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=L8RtQ14epmaddsttT155-n0VfNSjBWgTf3j2hR9s-zg';
    let PAURLRiskEMC ='https://prod-26.southeastasia.logic.azure.com:443/workflows/a7d6eb229b004a32aaa4f1010b5cdd0b/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=qr-4NGb7N9i-9AtZMTj4KW3UMlUjUVoWeEh9xgYCHRQ';

    let method='POST';
    
    switch (event.itemId) {
        //Risk Owner (Approved)
        case 'COMMAND_3':
              let userInput: string;
            
              userInput = await Dialog.prompt('Approve Remark :') ?? 'N/A';
              const raw = JSON.stringify({
                "MainID": listItemId,
                "Remarks": userInput,
                "Status": 'Approved',
                'Key': 'TaylorsUni'
              });
              this.callPowerAutomateFlowRiskOwner(PAURLRiskOwner,raw,method);

        // Show success message to the user
        await Dialog.alert('The remark has been approved successfully.');
        location.reload();

          break;
        //Risk Owner (Reject)
        case 'COMMAND_4':
            let userInputR: string;
            userInputR = await Dialog.prompt('Reject Remark : ') ?? 'N/A';
            const rawR = JSON.stringify({
              "MainID": listItemId,
              "Remarks": userInputR,
              "Status": 'Rejected',
              'Key': 'TaylorsUni'
            });
            this.callPowerAutomateFlowRiskOwner(PAURLRiskOwner,rawR,method);
        
        // Show success message to the user
        await Dialog.alert('The remark has been rejected successfully.');
        location.reload();

        break;

        //Risk Officer (Approved)
        case 'COMMAND_5':
          let userInputRo: string;
          userInputRo = await Dialog.prompt('Approved Remark : ') ?? 'N/A';
          const rawRo = JSON.stringify({
            "MainID": listItemId,
            "Remarks": userInputRo,
            "Status": 'Approved',
            'Key': 'TaylorsUni'
          });
          this.callPowerAutomateFlowRiskOfficer(PAURLRiskOfficer,rawRo,method);
// Show success message to the user
await Dialog.alert('The remark has been approved successfully.');
location.reload();

        break;
        //Risk Officer (Reject)
        case 'COMMAND_6':
          let userInputRoR: string;
          userInputRoR = await Dialog.prompt('Rejected Remark : ') ?? 'N/A';
          const rawRoR = JSON.stringify({
            "MainID": listItemId,
            "Remarks": userInputRoR,
            "Status": 'Rejected',
            'Key': 'TaylorsUni'
          });
          this.callPowerAutomateFlowRiskOfficer(PAURLRiskOfficer,rawRoR,method);
// Show success message to the user
await Dialog.alert('The remark has been rejected successfully.');
location.reload();
        break;
        
        //Risk RMOCS (Approved)
        case 'COMMAND_7':
          let userInputRMOC: string;
          userInputRMOC = await Dialog.prompt('Approved Remark : ') ?? 'N/A';
          const rawRMOC = JSON.stringify({
            "MainID": listItemId,
            "Remarks": userInputRMOC,
            "Status": 'Approved',
            'Key': 'TaylorsUni'
          });
          this.callPowerAutomateFlowRiskRMOC(PAURLRiskRMOC,rawRMOC,method);
// Show success message to the user
await Dialog.alert('The remark has been approved successfully.');
location.reload();

        break;
        //Risk RMOCS (Reject)
        case 'COMMAND_8':
          let userInputRMOCR: string;
          userInputRMOCR = await Dialog.prompt('Rejected Remark : ') ?? 'N/A';
          const rawRMOCR = JSON.stringify({
            "MainID": listItemId,
            "Remarks": userInputRMOCR,
            "Status": 'Rejected',
            'Key': 'TaylorsUni'
          });
          this.callPowerAutomateFlowRiskRMOC(PAURLRiskRMOC,rawRMOCR,method);
// Show success message to the user
await Dialog.alert('The remark has been rejected successfully.');
location.reload();
        break;

         //Risk EMC (Approved)
         case 'COMMAND_9':
          let userInputEMC: string;
          userInputEMC = await Dialog.prompt('Approved Remark : ') ?? 'N/A';
          const rawEMC = JSON.stringify({
            "MainID": listItemId,
            "Remarks": userInputEMC,
            "Status": 'Approved',
            'Key': 'TaylorsUni'
          });
          this.callPowerAutomateFlowRiskEMC(PAURLRiskEMC,rawEMC,method);
// Show success message to the user
await Dialog.alert('The remark has been approved successfully.');
location.reload();

        break;
        //Risk EMC (Reject)
        case 'COMMAND_10':
          let userInputEMCR: string;
          userInputEMCR = await Dialog.prompt('Rejected Remark : ') ?? 'N/A';
          const rawEMCR = JSON.stringify({
            "MainID": listItemId,
            "Remarks": userInputEMCR,
            "Status": 'Rejected',
            'Key': 'TaylorsUni'
          });
          this.callPowerAutomateFlowRiskEMC(PAURLRiskEMC,rawEMCR,method);
// Show success message to the user
await Dialog.alert('The remark has been rejected successfully.');
location.reload();
        break;


      default:
        throw new Error('Unknown command');
    }
  }

    //Call Power Automate Flow (Risk Owner)
    private async callPowerAutomateFlowRiskOwner( PAURL: string, JSON:string,method:string): Promise<void> {
      const requestOptions: RequestInit = {
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        },
        body:JSON,
        redirect:"follow",
        method:method
      };
      fetch(PAURL, requestOptions)
      .then((response) => response.text())
      .then((result) => console.log(result))
      .catch((error) => console.error(error));


  /* With bearer Token Code
    const provider =  await this.context.aadTokenProviderFactory.getTokenProvider();
    const token =   provider.getToken('https://taylorsedu.sharepoint.com');
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'Authorization': `bearer ${token}`,
      },
      body: JSON.stringify({
        "item": remarks
      })
    };
    
    
        this.context.spHttpClient.post('https://prod-46.southeastasia.logic.azure.com:443/workflows/5840ebbe6bbe48508e67d9350433b6b2/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Wu6cY-GiZmJlfrC0zbcd9HHuhQKlE9-JQL7_9468VFk', SPHttpClient.configurations.v1, options)
          .then((response: SPHttpClientResponse) => {
            if (response.ok) {
              console.log('Flow triggered successfully');
            } else {
              console.error(`Failed to trigger flow: ${response.statusText}`);
            }
          })
          .catch((error: any) => {
            console.error('Error triggering flow:', error);
          });
      
      */



    }
    //Call Power Automate Flow (Risk Officer)
    private async callPowerAutomateFlowRiskOfficer( PAURL: string, JSON:string,method:string): Promise<void> {
      const requestOptions: RequestInit = {
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        },
        body:JSON,
        redirect:"follow",
        method:method
      };
      fetch(PAURL, requestOptions)
      .then((response) => response.text())
      .then((result) => console.log(result))
      .catch((error) => console.error(error));





    }
    //Call Power Automate Flow (Risk RMOC)
    private async callPowerAutomateFlowRiskRMOC( PAURL: string, JSON:string,method:string): Promise<void> {
      const requestOptions: RequestInit = {
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        },
        body:JSON,
        redirect:"follow",
        method:method
      };
      fetch(PAURL, requestOptions)
      .then((response) => response.text())
      .then((result) => console.log(result))
      .catch((error) => console.error(error));
    }
    //Call Power Automate Flow (EMC)
    private async callPowerAutomateFlowRiskEMC( PAURL: string, JSON:string,method:string): Promise<void> {
          const requestOptions: RequestInit = {
            headers: {
              'Accept': 'application/json',
              'Content-Type': 'application/json'
            },
            body:JSON,
            redirect:"follow",
            method:method
          };
          fetch(PAURL, requestOptions)
          .then((response) => response.text())
          .then((result) => console.log(result))
          .catch((error) => console.error(error));
    }

    private async GetRiskOwnerFromDivisionDepartmentBreakdown():Promise<void>
    {
      try {
        const response: SPHttpClientResponse = await this.context.spHttpClient.get(`https://taylorsedu.sharepoint.com/sites/RiskManagement-STG/_api/web/lists/getbytitle('Division Department Breakdown')/items?$filter=RiskOwner/EMail eq '${this.context.pageContext.user.email}'`, SPHttpClient.configurations.v1);
        if (response.ok) {
            const data = await response.json();
            console.log('Risk Owner user:', data.value);
            if (data.value.length > 0) {
                this.properties.StateOwner = "RiskOwner";
            } else {
                this.properties.StateOwner = "0";
            }
        } else {
            console.error(`Failed to trigger flow: ${response.statusText}`);
            this.properties.StateOwner="0";
        }
    } catch (error) {
        console.error('Error triggering flow:', error);
        this.properties.StateOwner="0";
    }
    }
    private async GetRiskOfficerFromDivisionWiseApproval():Promise<void>
    {
      this.context.spHttpClient.get(`https://taylorsedu.sharepoint.com/sites/RiskManagement-STG/_api/web/lists/getbytitle('Division Wise Approval')/items?$filter=RiskOfficer/EMail eq '${this.context.pageContext.user.email}'`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
      //Risk Owner Button
              response.json().then((data) => {
                console.log('Risk Officer user:', data.value);
                // Process retrieved items here
                    if(data.value.length>0)
                      {
                        this.properties.State="RiskOfficer";
                      }else
                      {
                        this.properties.State="0";
                      }
        });
        } else {
            console.error(`Failed to trigger flow: ${response.statusText}`);
            this.properties.State="0";
          }
      })
      .catch((error: any) => {
        console.error('Error triggering flow:', error);
        this.properties.State="0";
      });


    }
      private async GetMultipleRMOCFromDivisionWiseApproval():Promise<void>
    {
      this.context.spHttpClient.get(`https://taylorsedu.sharepoint.com/sites/RiskManagement-STG/_api/web/lists/getbytitle('Division Wise Approval')/items?$filter=RMOCApprover/EMail eq '${this.context.pageContext.user.email}'`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
      //Risk Owner Button
              response.json().then((data) => {
                console.log('RMOC user:', data.value);
                // Process retrieved items here
                    if(data.value.length>0)
                      {
                        this.properties.StateRMOC="RiskRMOC";
                      }else
                      {
                        this.properties.StateRMOC="0";
                      }
        });
        } else {
            console.error(`Failed to trigger flow: ${response.statusText}`);
            this.properties.StateRMOC="0";
          }
      })
      .catch((error: any) => {
        console.error('Error triggering flow:', error);
        this.properties.StateRMOC="0";
      });


    }
    private async GetMultipleEMCFromDivisionWiseApproval():Promise<void>
    {
      this.context.spHttpClient.get(`https://taylorsedu.sharepoint.com/sites/RiskManagement-STG/_api/web/lists/getbytitle('Division Wise Approval')/items?$filter=EMCApprover/EMail eq '${this.context.pageContext.user.email}'`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
      //Risk Owner Button
              response.json().then((data) => {
                console.log('EMC user:', data.value);
                // Process retrieved items here
                    if(data.value.length>0)
                      {
                        this.properties.StateEMC="RiskEMC";
                      }else
                      {
                        this.properties.StateEMC="0";
                      }
        });
        } else {
            console.error(`Failed to trigger flow: ${response.statusText}`);
            this.properties.StateEMC="0";
          }
      })
      .catch((error: any) => {
        console.error('Error triggering flow:', error);
        this.properties.StateEMC="0";
      });


    }


  private _onListViewStateChanged =  async (args: ListViewStateChangedEventArgs): Promise<void> => {
      Log.info(LOG_SOURCE, 'List view state changed');

      // Retrieve the selected items and map them to the 'Risk_x0020_Statements' values
      let selectedItems = this.context.listView.selectedRows?.map(item => item.getValueByName('Risk_x0020_Statements')).toString();
      console.log("Get Risk Statement Value : " + selectedItems);

      console.log("Get ID : " + this.context.listView.selectedRows?.map(item => item.getValueByName('ID')).toString());


if (selectedItems) {
    // Regular expression to match a sequence of digits at the beginning of the string
    let numberMatch = selectedItems.match(/^\d+/);

    if (numberMatch) {
        // Convert the matched value to a number
        let number = parseInt(numberMatch[0], 10);
        console.log("Convert Number : " + numberMatch);
        

     const Statuss = await this.GetStatus(number);
        
 // TODO: Add your logic here
 const comparethreeCommand: Command = this.tryGetCommand('COMMAND_3');
 const comparefourCommand: Command = this.tryGetCommand('COMMAND_4');
 if(this.properties.StateOwner==="RiskOwner" && Statuss === "Submitted")
   {
     if (comparethreeCommand) {
       // This command should be hidden unless exactly one row is selected.
         comparethreeCommand.visible = this.context.listView.selectedRows?.length === 1;
     }
     if (comparefourCommand) {
       // This command should be hidden unless exactly one row is selected.
       comparefourCommand.visible = this.context.listView.selectedRows?.length === 1;
       }
 }

 const compare5Command: Command = this.tryGetCommand('COMMAND_5');
 const compare6Command: Command = this.tryGetCommand('COMMAND_6');
 if(this.properties.State==="RiskOfficer" && Statuss === "Pending at Risk Director Level")
   {
     if (compare5Command) {
       // This command should be hidden unless exactly one row is selected.
       compare5Command.visible = this.context.listView.selectedRows?.length === 1;
     }
     if (compare6Command) {
       // This command should be hidden unless exactly one row is selected.
       compare6Command.visible = this.context.listView.selectedRows?.length === 1;
       }
 }

 const compare7Command: Command = this.tryGetCommand('COMMAND_7');
 const compare8Command: Command = this.tryGetCommand('COMMAND_8');
 if(this.properties.StateRMOC === "RiskRMOC" && Statuss === "Pending at RMOC Level")
 {
   if (compare7Command) {
     // This command should be hidden unless exactly one row is selected.
     compare7Command.visible = this.context.listView.selectedRows?.length === 1;
   }
   if (compare8Command) {
     // This command should be hidden unless exactly one row is selected.
     compare8Command.visible = this.context.listView.selectedRows?.length === 1;
     }
 }

 const compare9Command: Command = this.tryGetCommand('COMMAND_9');
 const compare10Command: Command = this.tryGetCommand('COMMAND_10');
 if(this.properties.StateEMC === "RiskEMC" && Statuss === "Pending at EMC Level")
 {
   if (compare9Command) {
     // This command should be hidden unless exactly one row is selected.
     compare9Command.visible = this.context.listView.selectedRows?.length === 1;
   }
   if (compare10Command) {
     // This command should be hidden unless exactly one row is selected.
     compare10Command.visible = this.context.listView.selectedRows?.length === 1;
     }
 }

 console.log("Check Value");
 console.log("Raise ONchange");
 // You should call this.raiseOnChage() to update the command bar
this.raiseOnChange();
 
  }
}
else
{
//Hide All button
  const comparethreeCommand: Command = this.tryGetCommand('COMMAND_3');
const comparefourCommand: Command = this.tryGetCommand('COMMAND_4');
comparethreeCommand.visible = false;
comparefourCommand.visible = false;

 const compare5Command: Command = this.tryGetCommand('COMMAND_5');
 const compare6Command: Command = this.tryGetCommand('COMMAND_6');
 compare5Command.visible = false;
 compare6Command.visible = false;

 const compare7Command: Command = this.tryGetCommand('COMMAND_7');
 const compare8Command: Command = this.tryGetCommand('COMMAND_8');
 compare7Command.visible = false;
 compare8Command.visible = false;

 const compare9Command: Command = this.tryGetCommand('COMMAND_9');
 const compare10Command: Command = this.tryGetCommand('COMMAND_10');
 compare9Command.visible = false;
 compare10Command.visible = false;

  // You should call this.raiseOnChage() to update the command bar
 this.raiseOnChange();

}

  }


  private async GetStatus(number: number): Promise<string> {
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get( `https://taylorsedu.sharepoint.com/sites/RiskManagement-STG/_api/web/lists/getbytitle('RiskRegister')/items?$filter=ID eq '${number}'`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        console.log("Get Value Based on ID Response OK");

        const data = await response.json();
        if (data.value && data.value.length > 0) {
          console.log("Data:", data.value[0]);
          
          // Assuming 'Approver_x0020_Status' is the column name you are interested in
          let approvalStatus = data.value[0].Approver_x0020_Status;
          console.log('Approval Status:', approvalStatus);
          this.properties.Status = approvalStatus;

          // Process retrieved approval status here
          return approvalStatus;
        } else {
          console.log('No items found for the given ID');
          this.properties.Status = "";
          return "";
        }
      } else {
        console.log('Error fetching items:', response.statusText);
        return "";
      }
    } catch (error) {
      console.log('Error making the HTTP request:', error);
      return "";
    }
  }

/*

  private getCurrentUSer() :void {


    const comparethreeCommand: Command = this.tryGetCommand('COMMAND_3');

  this.context.spHttpClient.get(`https://taylorsedu.sharepoint.com/sites/RiskManagement-STG/_api/web/lists/getbytitle('Division Department Breakdown')/items?$filter=RiskOwner/EMail eq '${this.context.pageContext.user.email}'`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      if (response.ok) {
        comparethreeCommand.visible = true;

            response.json().then((data) => {
              console.log('Items for current user:', data.value);
              // Process retrieved items here
            });
  
    
        console.log('Flow triggered successfully');
      } else {
       
        comparethreeCommand.visible = false;
        console.error(`Failed to trigger flow: ${response.statusText}`);
      }
    })
    .catch((error: any) => {
      console.error('Error triggering flow:', error);
    
    });


    }
    */

}
