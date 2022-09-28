
import { SPFI } from "@pnp/sp";

export interface IListAPaloozaProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  sp: SPFI;

}

export interface IListAPaloozaState {
  items: any[];
}