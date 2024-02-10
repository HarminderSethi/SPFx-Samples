/* eslint-disable @typescript-eslint/no-explicit-any */
import { ServiceHealth } from "@microsoft/microsoft-graph-types";
import { MSGraphClientV3 } from "@microsoft/sp-http-msgraph";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export class GraphService {
  private context: WebPartContext;

  //private token: string;
  constructor(context: WebPartContext) {
    this.context = context;
  }

  private getClient = async (): Promise<MSGraphClientV3> => {
    return await this.context.msGraphClientFactory.getClient("3");
  };

  public getHealthOverviews = async (): Promise<ServiceHealth[]> => {
    const client = await this.getClient();
    const request = client.api("/admin/serviceAnnouncement/healthOverviews");
    const response = await request.expand("issues").get();
    return Promise.resolve(response.value);
  };
}
