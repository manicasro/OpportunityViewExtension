import { SPFI } from "@pnp/sp";
import styles from "../../view/ViewApplicationCustomizer.module.scss";
import { generatePickerWithButton } from "./generatePickerWithButtons";

/**
 * Function to generate an edit button for the opportunity
 */
export function generateEditButton(
  itemName: string,
  id: number,
  resetLastOpportunity: () => void,
  sp: SPFI,
  processOpportunity: () => Promise<void>
): HTMLElement {
  const editButton = document.createElement('button');
  editButton.className = styles.opportunityEditButton;
  editButton.innerHTML = "✎";
  editButton.addEventListener('click', () => {
    const targetDiv = document.getElementById(`${itemName}-value`);
    if (!!targetDiv) {
      const newDiv = generatePickerWithButton(itemName, id, targetDiv.innerText, resetLastOpportunity, sp, processOpportunity);
      targetDiv.replaceWith(newDiv);
    } else {
      console.log(`Div with id ${itemName}-name not found.`);
    }
  });
  return editButton;
}