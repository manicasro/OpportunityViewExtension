/* eslint-disable @typescript-eslint/no-explicit-any */
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

interface IUserInfo {
  Id: number;
  Title: string;
  Email: string;
}

/**
 * Function to get user information by user ID
 * @param userId 
 * @param spHttpClient 
 * @param context 
 * @returns 
 */
export async function getUserInfo(
  userId: string,
  spHttpClient: SPHttpClient,
  context: ApplicationCustomizerContext
): Promise<IUserInfo | undefined> {
  try {
    const response: SPHttpClientResponse = await spHttpClient.get(
      `${context.pageContext.web.absoluteUrl}/_api/web/getuserbyid(${userId})`,
      SPHttpClient.configurations.v1
    );

    if (response.ok) {
      const userInfo: IUserInfo = await response.json();
      return userInfo;
    } else {
      console.error(`Error getting user data: ${response.statusText}`);
      return undefined;
    }
  } catch (error) {
    console.error(`Exception while fetching user data: ${error}`);
    return undefined;
  }
}