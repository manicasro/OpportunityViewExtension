/**
 * Checks if the given URL contains the target site name.
 * @param url The current URL.
 * @param siteName The target site name to check for.
 * @returns True if the URL contains the site name, otherwise false.
 */
export function isOnTargetPage(url: string, siteName: string): boolean {
  return url.toLowerCase().indexOf(siteName.toLowerCase()) !== -1;
}