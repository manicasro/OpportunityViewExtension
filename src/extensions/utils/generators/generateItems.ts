import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { IOpportunity } from "../../../IOpportunity";
import styles from "../../view/ViewApplicationCustomizer.module.scss";
import { getUserInfo } from "../apiCalls/getUserInfo";
import { generateDateItem } from "./generateDateItem";
import { generateOpportunityItem } from "./generateOpportunityItem";
import { SPHttpClient } from '@microsoft/sp-http';
import { SPFI } from "@pnp/sp";

export async function generateItems(
  data: IOpportunity,
  spHttpClient: SPHttpClient,
  context: ApplicationCustomizerContext,
  resetLastOpportunity: () => void,
  sp: SPFI,
  processOpportunity: () => Promise<void>
): Promise<HTMLElement> {
  const divElem = document.createElement('div');
  divElem.className = styles.opportunityItems;

  try {
      // Fetch user information and generate date items concurrently
      const results = await Promise.all([
          data.sfaSalerStringId ? getUserInfo(data.sfaSalerStringId, spHttpClient, context) : null,
          data.sfaBidManagerStringId ? getUserInfo(data.sfaBidManagerStringId, spHttpClient, context) : null,
          data.sfaGarantStringId ? getUserInfo(data.sfaGarantStringId, spHttpClient, context) : null,
          data.sfaLegalStringId ? getUserInfo(data.sfaLegalStringId, spHttpClient, context) : null,
          data.sfaTechnicalGarantStringId ? getUserInfo(data.sfaTechnicalGarantStringId, spHttpClient, context) : null,
          generateDateItem('Termín Vysvětlení', data.sfaExplanationDate, 'sfaExplanationDate', data.Id, spHttpClient, context, resetLastOpportunity, sp, processOpportunity),
          generateDateItem('Termín ÚOHS', data.sfaUohsDate, 'sfaUohsDate', data.Id, spHttpClient, context, resetLastOpportunity, sp, processOpportunity),
      ]);

      // Extract user names from results
      const [saler, bidManager, garant, legal, technicalGarant, explanationDate, uohsDate] = results;

      const salerName = saler?.Title || "";
      const managerName = bidManager?.Title || "";
      const garantName = garant?.Title || "";
      const legalName = legal?.Title || "";
      const technicalGarantName = technicalGarant?.Title || "";

      // Append opportunity items to the container
      divElem.appendChild(generateOpportunityItem('Zadavatel', data.sfaCustomer));
      divElem.appendChild(generateOpportunityItem('Status VZ', data.sfaGoNoGo));
      divElem.appendChild(generateOpportunityItem('RFP Day', data.sfaRfpDay));
      divElem.appendChild(generateOpportunityItem('Obchodník', salerName));
      divElem.appendChild(generateOpportunityItem('Garant nabídky', garantName));
      divElem.appendChild(generateOpportunityItem('BID manažer', managerName));
      divElem.appendChild(generateOpportunityItem('Právní konzultant', legalName));
      divElem.appendChild(generateOpportunityItem('Technický Garant', technicalGarantName));
      divElem.appendChild(generateOpportunityItem('Fáze příležitosti', data.sfaOpportunityPhase));
      divElem.appendChild(generateOpportunityItem('Důvod prohry', data.sfaReasonOfLost));

      // Append date items
      divElem.appendChild(explanationDate as HTMLElement);
      divElem.appendChild(uohsDate as HTMLElement);
  } catch (error) {
      console.error('Error generating items:', error);
  }

  return divElem;
}