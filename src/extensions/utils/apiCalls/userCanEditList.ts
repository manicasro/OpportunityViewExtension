import { SPPermission } from '@microsoft/sp-page-context';
import { SPHttpClient } from '@microsoft/sp-http';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';

/**
 * Function to check if the user has permission to edit the list
 */
export async function userCanEditList(spHttpClient: SPHttpClient, context: ApplicationCustomizerContext): Promise<boolean> {
  try {
    const response = await spHttpClient.get(`${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('oneSfaRecordsList')/EffectiveBasePermissions`, SPHttpClient.configurations.v1);
    const permissions = await response.json();
    const manageListsPermission: SPPermission = new SPPermission(permissions);
    return manageListsPermission.hasPermission(SPPermission.manageLists);
  } catch (error) {
    console.error(error);
    return false;
  }
}