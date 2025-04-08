import { SPFI } from "@pnp/sp";
import styles from "../../view/ViewApplicationCustomizer.module.scss";
import { updateSfaExplanationDate } from "../apiCalls/updateSfaExplanationDate";

export function generateConfirmButton(
  datePicker: HTMLInputElement,
  id: number,
  itemName: string,
  resetLastOpportunity: () => void,
  sp: SPFI,
  processOpportunity: () => Promise<void>
): HTMLButtonElement {
  const confirmButton = document.createElement('button');
  confirmButton.className = styles.opportunityConfirmButton;
  confirmButton.innerHTML = '\u2713';
  confirmButton.addEventListener('click', async () => {
      try {
          const selectedDate = datePicker.value;
          await updateSfaExplanationDate(resetLastOpportunity, id, selectedDate, sp, itemName, processOpportunity);
      } catch (error) {
          console.error('Error processing opportunity:', error);
      }
  });
  return confirmButton;
}