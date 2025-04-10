import styles from "../../view/ViewApplicationCustomizer.module.scss";
import { userCanEditList } from "../apiCalls/userCanEditList";
import { generateEditableDateDiv } from "./generateEditableDiv";
import { generatePickerWithButton } from "./generatePickerWithButtons";
import { SPHttpClient } from '@microsoft/sp-http';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { SPFI } from "@pnp/sp";

/**
 * Function to generate a date item view for an opportunity
 */
export async function generateDateItem(
  parameterName: string,
  parameterValue: string,
  itemName: string,
  id: number,
  spHttpClient: SPHttpClient,
  context: ApplicationCustomizerContext,
  resetLastOpportunity: () => void,
  sp: SPFI,
  processOpportunity: () => Promise<void>
): Promise<HTMLElement> {
  const divElem = document.createElement('div');
  divElem.className = styles.opportunityDateItemView;

  const name = document.createElement('p');
  name.className = styles.opportunityItemParamName;
  name.innerHTML = parameterName;
  divElem.appendChild(name);

  const val = document.createElement('p');
  val.className = styles.opportunityItemParamValue;
  if (!parameterValue) {
    userCanEditList(spHttpClient, context).then((canEdit) => {
      if (!canEdit) {
        val.innerHTML = 'N/A';
        divElem.appendChild(val);
        return divElem;
      }else{
        divElem.appendChild(generatePickerWithButton(itemName, id, '', resetLastOpportunity, sp, processOpportunity));
        return divElem;
      }
    })
    .catch((error) => {
      console.error('Error checking user permissions:', error);
    });
  } else {
    divElem.appendChild(generateEditableDateDiv(parameterValue, itemName, id, resetLastOpportunity, sp, processOpportunity));
    return divElem;
  }
  return divElem;
}