/* eslint-disable */

import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IOpportunity } from '../../IOpportunity';
import styles from './ViewApplicationCustomizer.module.scss';
import { SPPermission } from '@microsoft/sp-page-context';
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/behaviors/spfx";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { DefaultConfig, IConfig } from '../config/Config';
import { isOnTargetPage } from '../utils/UrlUtils';
import { PollingService } from '../PollingService';


export interface IViewApplicationCustomizerProperties {
  testMessage: string;
}

export default class ViewApplicationCustomizer
  extends BaseApplicationCustomizer<IViewApplicationCustomizerProperties> {

  private spHttpClient: SPHttpClient;
  private config: IConfig = DefaultConfig;       
  private previousUrl: string;
  private currentlyOnSiteWithoutInfo : boolean = false;
  private lastOpportunity: string = '';
  private urlPollingIntervalId: number | null = null;
  private sp: SPFI;
  private pollingService: PollingService = new PollingService();

  public async onInit(): Promise<void> {
    console.log("Initializing ViewApplicationCustomizer extension.");

    // Initialize services (PnP, SPHttpClient, etc.)
    this.initializeServices();

    // Start polling for URL changes
    this.startUrlPolling();

    return Promise.resolve();
  }

  private initializeServices(): void {
    console.log("Initializing services...");

    // Initialize PnP JS
    this.sp = spfi().using(SPFx(this.context));

    // Obtain SPHttpClient instance from context
    this.spHttpClient = this.context.spHttpClient;

    // Save the initial URL
    this.previousUrl = window.location.href;

    console.log("Services initialized.");
  }

  protected onDispose(): void {
    // Remove the custom div when the extension is disposed
    if (!!this.urlPollingIntervalId) {
      clearInterval(this.urlPollingIntervalId);
    }
  }

  private startUrlPolling(): void {
    console.log("Starting URL polling.");
  
    // Use the PollingService to manage the interval
    this.pollingService.startPolling(() => {
      this.checkForUrlChange();
    }, 500); // Poll every half second (adjust interval as needed)
  }

  private checkForUrlChange(): void {
    const currentUrl = window.location.href;
  
    // Check if the URL has changed
    if (currentUrl !== this.previousUrl) {
      this.previousUrl = currentUrl;
      this.currentlyOnSiteWithoutInfo = false;
    }
  
    // Handle URL-specific logic
    if (isOnTargetPage(currentUrl, this.config.siteName)) {
      this.processOpportunity();
    } else {
      this.removeInjectedExtensionDiv();
    }
  }

  private removeInjectedExtensionDiv(): void {
    const divToRemove = document.getElementById("InjectedExtensionDiv");
    if (divToRemove?.parentNode) {
      divToRemove.parentNode.removeChild(divToRemove);
      console.log("Injected extension div removed.");
    }
  
    // Reset the last opportunity to undefined for clarity
    this.lastOpportunity = '';
  }

  private async processOpportunity(): Promise<void> {
    let opportunity: string | null = this.parseUrl();
    // If opportunity is not found, remove the injected div and return
    if (!opportunity) {
      this.removeInjectedExtensionDiv();
      return;
    }

    // Log if opportunity has changed
    if (opportunity !== this.lastOpportunity) {
      console.log(`Opportunity changed - ${opportunity}. Fetching new data.`);
    }

    // Check if the injected div already exists
    let injectedDiv = document.getElementById("InjectedExtensionDiv");
    
    if (!injectedDiv) {
      if (this.currentlyOnSiteWithoutInfo) {
        this.lastOpportunity = opportunity;
        return Promise.resolve();
      }
      const data = await this.fetchData(opportunity);
      if (!!data) {
        this.renderCustomDiv(data);
      } else {
        this.currentlyOnSiteWithoutInfo = true;
      }
    } else {
      if (this.lastOpportunity !== opportunity) {
        const data = await this.fetchData(opportunity);
        if (!!data) {
          this.renderCustomDiv(data);
        } else {
          this.removeInjectedExtensionDiv();
        }
      }
    }
    this.lastOpportunity = opportunity;
    return Promise.resolve();
  }

  private renderCustomDiv(data: IOpportunity): void {
    // Create or update the dynamic content
    let injectedDiv = document.getElementById("InjectedExtensionDiv");

      // Dynamically adjust the grid structure
      const mainContainer = document.querySelector('.main_9c0f266f') as HTMLElement;
      if (mainContainer) {
          // Update grid-template-areas
          mainContainer.style.gridTemplateAreas = `
              "spfxHeader spfxHeader"
              "commandBar commandBar"
              "renderAfterCommandBar renderAfterCommandBar"
              "injectedDiv injectedDiv"
              "messageBar messageBar"
              "header pane"
              "headerBar pane"
              "contentBar pane"
              "content pane"
              "spfxFooter spfxFooter"
              "debug debug"
          `;

          // Update grid-template-rows
          mainContainer.style.gridTemplateRows = `
              max-content max-content max-content max-content max-content max-content max-content max-content 2fr max-content auto
          `;
      }
      // Remove the existing div if it exists
      if (injectedDiv) {
        injectedDiv.parentNode?.removeChild(injectedDiv);
      }
      // Insert the new div
      mainContainer.appendChild(this.generateInjectedDiv(data));
  }
   

  private parseUrl(): string | null {
    // Find URL, parse it and call the correct endpoint with REST API
    const url = window.location.href;
    const decodedUrl = decodeURIComponent(url);
    // Find the index of the part that starts with keySequence
    const keySequence = this.config.keySequence;
    let idIndex;

    for (let str of keySequence) {
      idIndex = decodedUrl.indexOf(str);
      if (idIndex !== -1) {
        const partsAfterId = decodedUrl.substring(idIndex + str.length).split('/');
        if (partsAfterId.length < 4) {
          return null;
        } else {
          return partsAfterId[3].split('&')[0];
        }
      }
    }
    return null;
  }

  private async fetchData(opportunity: string): Promise<IOpportunity | null> {
    try {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('oneSfaRecordsList')/items?$filter=sfaLeadId eq '${opportunity}'`;
      const response: SPHttpClientResponse = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
  
      if (!response.ok) {
        console.error(`Failed to fetch data. Status: ${response.status}, StatusText: ${response.statusText}`);
        return null;
      }
  
      const data = await response.json();
      console.log("Json received", data);
  
      // Return the first item if data is found
      return data.value && data.value.length > 0 ? (data.value[0] as IOpportunity) : null;
    } catch (error) {
      console.error("Error fetching data:", error);
      return null;
    }
  }

  // Method to fetch user information by ID
  private getUserInfo(userId: string): Promise<any> {
    return this.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/getuserbyid(${userId})`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          console.error(`Error getting user data: ${response.statusText}`);
          return Promise.reject(response.statusText);
        }
      });
  }

  // Method to generate Opportunity items
  private generateOpportunityItem(parameterName: string, parameterValue: string): HTMLElement {
    if (parameterValue === null || parameterValue === undefined) {
      parameterValue = 'N/A';
    }
    
    let divElem = document.createElement('div');
    divElem.className = styles.opportunityItemView;

    let parameterNamePar = document.createElement('p');
    parameterNamePar.className = styles.opportunityItemParamName;
    parameterNamePar.innerHTML = parameterName;

    let parameterValuePar = document.createElement('p');
    parameterValuePar.className = styles.opportunityItemParamValue;
    parameterValuePar.innerHTML = parameterValue;

    divElem.appendChild(parameterNamePar);
    divElem.appendChild(parameterValuePar);

    return divElem;
  }

  private generateButtons(data: IOpportunity): HTMLElement {
    let divElem = document.createElement('div');
    divElem.className = styles.opportunityButtonContainer;


    let teamsButton : HTMLButtonElement | null = document.createElement('button');

    if ((data.sfaTeamId !== null && data.sfaTeamId !== undefined) &&
        (data.sfaGenChannel !== null && data.sfaGenChannel !== undefined) &&
        (this.config.tenantId !== null && this.config.tenantId !== undefined) &&
        !!data.sfaTeamDone) {
      teamsButton.className = styles.opportunityLinkButton;
      teamsButton.innerHTML = 'Teams';
      teamsButton.addEventListener('click', () => {
        const sfaGenChannel = data.sfaGenChannel;
        const sfaTeamId = data.sfaTeamId;
        const teamsUrl = `https://teams.microsoft.com/l/channel/${sfaGenChannel}/General?groupId=${sfaTeamId}&tenantId=${this.config.tenantId}`;
        window.open(teamsUrl, '_blank');
      });
    }else{
      teamsButton = null;
    }

    let salesForceButton: HTMLButtonElement | null = document.createElement('button');
    let salesForceUrl: string;
    if (data.sfaOpportunityId !== null && data.sfaOpportunityId !== undefined &&
        this.config.opportunityUrl !== null && this.config.opportunityUrl !== undefined) {
      salesForceUrl = `${this.config.opportunityUrl}${data.sfaOpportunityId}`;
    }else if (data.sfaLeadId !== null && data.sfaLeadId !== undefined &&
              this.config.leadUrl !== null && this.config.leadUrl !== undefined) {
      salesForceUrl = `${this.config.leadUrl}${data.sfaLeadId}`;
    }else{
      salesForceButton = null;
    }

    if (teamsButton !== null) {
      divElem.appendChild(teamsButton);
    }
    if (salesForceButton !== null) {
      salesForceButton.className = styles.opportunityLinkButton;
      salesForceButton.innerHTML = 'SalesForce';
      salesForceButton.addEventListener('click', () => {
        window.open(salesForceUrl, '_blank');
      });
      divElem.appendChild(salesForceButton);
    }
    
    return divElem;
  }

  private generateDatePicker(itemName: string, date: string | null): HTMLInputElement {
    let datePicker: HTMLInputElement;
    datePicker = document.createElement('input');
    datePicker.type = 'date';
    datePicker.className = styles.opportunityDatePicker;
    datePicker.id = `${itemName}-datePicker`; // add an id to the input
    if (date !== '' && date !== null && date !== undefined) {
      // Split the date string and parse day, month, year
      let [day, month, year] = date.split('.');
      // Create a new Date object using year, month, day (month - 1 because months are zero-indexed)
      let dateObj = new Date(parseInt(year), parseInt(month)-1, parseInt(day)+1);
      // Convert the date object to ISO string and set as value
      datePicker.value = dateObj.toISOString().slice(0,10);
    }
    return datePicker;
  }

  private async updateSfaExplanationDate(Id: number, selectedDate: string, itemName: string): Promise<void> {
    if (selectedDate === '' || selectedDate === null || selectedDate === undefined) return;
    const formattedDate = `${selectedDate}T00:00:00Z`; // format the date in ISO 8601 format
    if (itemName === 'sfaExplanationDate') {
      this.sp.web.lists.getByTitle('oneSfaRecordsList').items.getById(Id).update({ sfaExplanationDate: formattedDate }).then(() => {
        this.lastOpportunity = '';
        this.processOpportunity();});
    } else {
      this.sp.web.lists.getByTitle('oneSfaRecordsList').items.getById(Id).update({ sfaUohsDate: formattedDate }).then(() => {
        this.lastOpportunity = '';
        this.processOpportunity();});
    }
  }

  private generateConfirmButton(datePicker: HTMLInputElement, id: number, itemName: string): HTMLButtonElement {
    let confirmButton = document.createElement('button');
    confirmButton.className = styles.opportunityConfirmButton;
    confirmButton.innerHTML = '\u2713';
    confirmButton.addEventListener('click', () => {
      let selectedDate = datePicker.value; 
      //console.log(`${selectedDate} - ${id}`); // logs the selected date in 'yyyy-mm-dd' format
      this.updateSfaExplanationDate(id, selectedDate, itemName);
    });
    return confirmButton;
  }

  private generatePickerWithButton(itemName: string, id: number, date: string): HTMLElement {
    let divElem = document.createElement('div');
    divElem.className = styles.opportunityPickerAndButton;

    let datePicker = this.generateDatePicker(itemName, date);
    let confirmButton = this.generateConfirmButton(datePicker, id, itemName);

    divElem.appendChild(datePicker);
    divElem.appendChild(confirmButton);

    return divElem;
  }

  generateEditButton(itemName: string, id: number): HTMLElement {
    let editButton = document.createElement('button');
    editButton.className = styles.opportunityEditButton;
    editButton.innerHTML = "✎";
    editButton.addEventListener('click', () => {
      let targetDiv = document.getElementById(`${itemName}-value`);
      if (!!targetDiv) {
        //console.log("Button was clicked and the former date we want to edit is: ", targetDiv.innerText);
        let newDiv = this.generatePickerWithButton(itemName, id, targetDiv.innerText);
        targetDiv.replaceWith(newDiv);
      } else {
        console.log(`Div with id ${itemName}-name not found.`);
      }
    });
    return editButton;
  }

  private generateEditableDateDiv(parameterValue: string, itemName: string, id: number): HTMLElement {
    let divElem = document.createElement('div');
    divElem.id = `${itemName}-value`;
    divElem.className = styles.opportunityEditableDate;
  
    let val = document.createElement('p');
    val.className = styles.opportunityItemParamValue;
    const date: Date = new Date(parameterValue);  
    val.innerHTML = date.getDate() + '.' + (date.getMonth() + 1) + '.' + date.getFullYear();
    divElem.appendChild(val);
    divElem.appendChild(this.generateEditButton(itemName, id));
   
    return divElem;
  }

  private async userCanEditList(): Promise<boolean> {
    try {
      const response = await this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('oneSfaRecordsList')/EffectiveBasePermissions`, SPHttpClient.configurations.v1);
      const permissions = await response.json();
      const manageListsPermission: SPPermission = new SPPermission(permissions);
      return manageListsPermission.hasPermission(SPPermission.manageLists);
    } catch (error) {
      console.error(error);
      return false;
    }
  }

  private async generateDateItem(parameterName: string, parameterValue: string, itemName: string, id: number): Promise<HTMLElement> {
    let divElem = document.createElement('div');
    divElem.className = styles.opportunityDateItemView;
  
    let name = document.createElement('p');
    name.className = styles.opportunityItemParamName;
    name.innerHTML = parameterName;
    divElem.appendChild(name);
  
    let val = document.createElement('p');
    val.className = styles.opportunityItemParamValue;
    if (parameterValue === null || parameterValue === undefined) {
      this.userCanEditList().then((canEdit) => {
        if (!canEdit) {
          val.innerHTML = 'N/A';
          divElem.appendChild(val);
          return divElem;
        }else{
          divElem.appendChild(this.generatePickerWithButton(itemName, id, ''));
          return divElem;
        }
      });
    } else {
      divElem.appendChild(this.generateEditableDateDiv(parameterValue, itemName, id));
      return divElem;
    }
    return divElem;
  }

  private generateItems(data: IOpportunity): HTMLElement {
    let divElem = document.createElement('div');
    divElem.className = styles.opportunityItems;

    // Fetch user information for each ID
    Promise.all([
      (data.sfaSalerStringId === null || data.sfaSalerStringId == undefined) 
        ? null 
        : this.getUserInfo(data.sfaSalerStringId),
      (data.sfaBidManagerStringId === null || data.sfaBidManagerStringId == undefined) 
        ? null 
        : this.getUserInfo(data.sfaBidManagerStringId),
      (data.sfaGarantStringId === null || data.sfaGarantStringId == undefined) 
        ? null 
        : this.getUserInfo(data.sfaGarantStringId),
      (data.sfaLegalStringId === null || data.sfaLegalStringId == undefined) 
        ? null 
        : this.getUserInfo(data.sfaLegalStringId),
      (data.sfaTechnicalGarantStringId === null || data.sfaTechnicalGarantStringId == undefined) 
        ? null 
        : this.getUserInfo(data.sfaTechnicalGarantStringId),
      this.generateDateItem('Termín Vysvětlení', data.sfaExplanationDate, 'sfaExplanationDate', data.Id),
      this.generateDateItem('Termín ÚOHS', data.sfaUohsDate, 'sfaUohsDate', data.Id)
    ])
    .then((arr: any[]) => {
      const salerName = (arr[0] === null || arr[0] == undefined)
      ? null
      : arr[0].Title;
      const managerName = (arr[1] === null || arr[1] == undefined)
      ? null
      : arr[1].Title;
      const garantName = (arr[2] === null || arr[2] == undefined)
      ? null
      : arr[2].Title;
      const legalName = (arr[3] === null || arr[3] == undefined)
      ? null
      : arr[3].Title;
      const technicalGarantName = (arr[4] === null || arr[4] == undefined)
      ? null
      : arr[4].Title;
      
      divElem.appendChild(this.generateOpportunityItem('Zadavatel', data.sfaCustomer));
      divElem.appendChild(this.generateOpportunityItem('Status VZ', data.sfaGoNoGo));
      divElem.appendChild(this.generateOpportunityItem('RFP Day', data.sfaRfpDay));
      divElem.appendChild(this.generateOpportunityItem('Obchodník', salerName));
      divElem.appendChild(this.generateOpportunityItem('Garant nabídky', garantName));
      divElem.appendChild(this.generateOpportunityItem('BID manažer', managerName));
      divElem.appendChild(this.generateOpportunityItem('Právní konzultant', legalName));
      divElem.appendChild(this.generateOpportunityItem('Technický Garant', technicalGarantName));
      divElem.appendChild(this.generateOpportunityItem('Fáze příležitosti', data.sfaOpportunityPhase));
      divElem.appendChild(this.generateOpportunityItem('Důvod prohry', data.sfaReasonOfLost));
      
      divElem.appendChild(arr[5]);
      divElem.appendChild(arr[6]);
    })

    return divElem;
  }

  private generateTitle(title: string): HTMLElement {
    let divElem = document.createElement('div');
    divElem.className = styles.opportunityTitleContainer;

    let titleParam = document.createElement('p');
    titleParam.className = styles.opportunityTitleParam;
    titleParam.innerHTML = 'Název zakázky';

    let titleValue = document.createElement('p');
    titleValue.className = styles.opportunityTitleValue;
    titleValue.innerHTML = title;

    divElem.appendChild(titleParam);
    divElem.appendChild(titleValue);

    return divElem;
  }

  private generateContent(data: IOpportunity): HTMLElement {
    let divElem = document.createElement('div');
    divElem.className = styles.opportunityViewContent;
    
    let titleDiv = this.generateTitle(data.sfaLeadName);
    let itemsDiv = this.generateButtons(data);

    divElem.appendChild(titleDiv);
    divElem.appendChild(itemsDiv);

    return divElem;
  }

  private generateInjectedDiv(data: IOpportunity): HTMLElement {
    const wholeDiv = document.createElement("div");
    wholeDiv.className = styles.wholeDiv;

    const baseDiv = document.createElement("div");

    baseDiv.setAttribute("id", "InjectedExtensionDiv");
    baseDiv.className = styles.baseInjectedDiv

    wholeDiv.appendChild(this.generateContent(data));
    wholeDiv.appendChild(this.generateItems(data));

    baseDiv.appendChild(wholeDiv);

    return baseDiv;
  }
}
/* eslint-enable */