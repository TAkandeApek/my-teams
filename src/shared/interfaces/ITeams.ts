export interface ITeam {
    id: string;
    displayName: string;
    description: string;
    isArchived: boolean;  
    visibility: string; 
    imagePath: string;  
    channelsData: IChannel[];
    createdDateTime: Date;
    resourceProvisioningOptions: string[]
  }

  export interface ITeamFinal {
    id: string;
    displayName: string;
    description: string;
    isArchived: boolean;
    teamImage: string;
  }

  export interface ITeamImage {
    content: string;
  }

  export interface IChannel {
    id: string;
    displayName: string;
    description: string;
    webUrl: string;
  }
  