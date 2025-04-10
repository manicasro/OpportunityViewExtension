/* eslint-disable */

import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IOpportunity } from '../../IOpportunity';
import styles from './ViewApplicationCustomizer.module.scss';
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/behaviors/spfx";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { DefaultConfig, IConfig } from '../config/Config';
import { isOnTargetPage } from '../utils/UrlUtils';
import { PollingService } from '../PollingService';
import { generateContent } from '../utils/generators/generateContent';
import { generateItems } from '../utils/generators/generateItems';


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

  private async renderCustomDiv(data: IOpportunity): Promise<void> {
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
      
      // Generate the new injected div
      const newInjectedDiv = await this.generateInjectedDiv(data);

      // Replace the existing div if it exists, otherwise append the new one
      if (injectedDiv) {
          injectedDiv.replaceWith(newInjectedDiv);
      } else {
          mainContainer.appendChild(newInjectedDiv);
      }
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

  private resetLastOpportunity(): void {
    this.lastOpportunity = '';
  }

  private async generateInjectedDiv(data: IOpportunity): Promise<HTMLElement> {
    const wholeDiv = document.createElement("div");
    wholeDiv.className = styles.wholeDiv;

    const baseDiv = document.createElement("div");
    baseDiv.setAttribute("id", "InjectedExtensionDiv");
    baseDiv.className = styles.baseInjectedDiv

    wholeDiv.appendChild(generateContent(data, this.config));

    // Await the result of generateItems
    const itemsDiv = await generateItems(
        data,
        this.spHttpClient,
        this.context,
        this.resetLastOpportunity.bind(this),
        this.sp,
        this.processOpportunity.bind(this)
    );
    wholeDiv.appendChild(itemsDiv);
    
    baseDiv.appendChild(wholeDiv);

    return baseDiv;
  }
}
/* eslint-enable */