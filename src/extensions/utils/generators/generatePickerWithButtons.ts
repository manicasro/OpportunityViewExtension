import { SPFI } from "@pnp/sp";
import styles from "../../view/ViewApplicationCustomizer.module.scss";
import { generateConfirmButton } from "./generateConfirmButton";
import { generateDatePicker } from "./generateDatePicker";

export function generatePickerWithButton(
  itemName: string,
  id: number,
  date: string,
  resetLastOpportunity: () => void,
  sp: SPFI,
  processOpportunity: () => Promise<void>
): HTMLElement {
  const divElem = document.createElement('div');
  divElem.className = styles.opportunityPickerAndButton;

  const datePicker = generateDatePicker(itemName, date);
  const confirmButton = generateConfirmButton(datePicker, id, itemName, resetLastOpportunity, sp, processOpportunity);

  divElem.appendChild(datePicker);
  divElem.appendChild(confirmButton);

  return divElem;
}