import styles from "../../view/ViewApplicationCustomizer.module.scss";

/**
 * Function to generate a date picker input element
 * @param itemName 
 * @param date 
 * @returns 
 */
export function generateDatePicker(itemName: string, date: string | undefined): HTMLInputElement {
    const datePicker = document.createElement('input');
    datePicker.type = 'date';
    datePicker.className = styles.opportunityDatePicker;
    datePicker.id = `${itemName}-datePicker`; // add an id to the input
    if (!!date) {
      // Split the date string and parse day, month, year
      const [day, month, year] = date.split('.');
      // Create a new Date object using year, month, day (month - 1 because months are zero-indexed)
      const dateObj = new Date(parseInt(year), parseInt(month)-1, parseInt(day)+1);
      // Convert the date object to ISO string and set as value
      datePicker.value = dateObj.toISOString().slice(0,10);
    }
    return datePicker;
  }