import { IOpportunity } from "../../../IOpportunity";
import { IConfig } from "../../config/Config";
import styles from "../../view/ViewApplicationCustomizer.module.scss";
import { generateButtons } from "./generateButtons";
import { generateTitle } from "./generateTitle";

/**
 * Function to generate the content of the opportunity view
 * @param data 
 * @param config 
 * @returns 
 */
export function generateContent(data: IOpportunity, config: IConfig): HTMLElement {
  const divElem = document.createElement('div');
  divElem.className = styles.opportunityViewContent;
  
  const titleDiv = generateTitle(data.sfaLeadName);
  const itemsDiv = generateButtons(data, config);

  divElem.appendChild(titleDiv);
  divElem.appendChild(itemsDiv);

  return divElem;
}