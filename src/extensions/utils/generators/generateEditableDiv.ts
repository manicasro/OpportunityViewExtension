import { SPFI } from "@pnp/sp";
import styles from "../../view/ViewApplicationCustomizer.module.scss";
import { generateEditButton } from "./generateEditButton";

/**
 * Function to generate an editable date div for the opportunity
 */
export function generateEditableDateDiv(
  parameterValue: string,
  itemName: string,
  id: number,
  resetLastOpportunity: () => void,
  sp: SPFI,
  processOpportunity: () => Promise<void>
): HTMLElement {
    const divElem = document.createElement('div');
    divElem.id = `${itemName}-value`;
    divElem.className = styles.opportunityEditableDate;
  
    const val = document.createElement('p');
    val.className = styles.opportunityItemParamValue;
    const date: Date = new Date(parameterValue);  
    val.innerHTML = date.getDate() + '.' + (date.getMonth() + 1) + '.' + date.getFullYear();
    divElem.appendChild(val);
    divElem.appendChild(generateEditButton(itemName, id, resetLastOpportunity, sp, processOpportunity));
   
    return divElem;
  }