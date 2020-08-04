import { ITeam, IChannel,ITeamImage } from "../interfaces/ITeams";

export interface ITeamsService {
  GetTeams(): Promise<ITeam[]>;
  GetTeamChannels(teamId): Promise<IChannel[]>;
  GetTeamPhoto(teamId): Promise<string>;
}
