export interface IConfig {
  tenantId: string;
  opportunityUrl: string;
  leadUrl: string;
  siteName: string;
  keySequence: string[];
}

export const DefaultConfig: IConfig = {
  tenantId: "b213b057-1008-4204-8c53-8147bc602a29",
  opportunityUrl: "https://tmobileczsk--situat.sandbox.lightning.force.com/lightning/cmp/coredt__NavigateTo?c__objectName=Opportunity&c__externalId=",
  leadUrl: "https://tmobileczsk--situat.sandbox.lightning.force.com/lightning/cmp/coredt__NavigateTo?c__objectName=Lead&c__externalId=",
  siteName: "sites/f-test-zakazky/verejne_zakazky",
  keySequence: ["id=/", "RootFolder=/"],
};