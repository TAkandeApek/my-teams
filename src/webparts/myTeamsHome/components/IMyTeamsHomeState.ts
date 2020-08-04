import { ITeam, IChannel } from "../../../shared/interfaces";

export interface IMyTeamsHomeState {
  items: ITeam[]; 
  SortedBy: string; 
}
