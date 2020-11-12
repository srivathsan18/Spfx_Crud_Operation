import { SPHttpClient } from '@microsoft/sp-http'
export interface ITestProps {
  description: string;
  spHttpClient: SPHttpClient;  
  siteUrl: string;  
  listName:string;
}
