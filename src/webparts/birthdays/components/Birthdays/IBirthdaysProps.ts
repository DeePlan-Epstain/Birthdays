import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";

export interface IBirthdaysProps {
  description: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  BirthdayListId: string;
  sp: SPFI;
  BirthdaysRange: string;
  context: WebPartContext;
  BackgroundImage: string;
  pageItemsNumber: number;

  SpotlightId: string;
  SpotlightInterval: number;
}
