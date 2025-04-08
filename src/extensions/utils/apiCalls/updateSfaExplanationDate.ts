import { SPFI } from "@pnp/sp";

export async function updateSfaExplanationDate(
  resetLastOpportunity: () => void,
  Id: number,
  selectedDate: string,
  sp: SPFI,
  itemName: string,
  processOpportunity: () => Promise<void>
): Promise<void> {
  if (!selectedDate) return;

  const formattedDate = `${selectedDate}T00:00:00Z`; // format the date in ISO 8601 format
  const fieldToUpdate = itemName === 'sfaExplanationDate' ? { sfaExplanationDate: formattedDate } : { sfaUohsDate: formattedDate };

  await sp.web.lists.getByTitle('oneSfaRecordsList').items.getById(Id).update(fieldToUpdate);
  resetLastOpportunity(); // Reset the last opportunity after updating
  await processOpportunity();
}