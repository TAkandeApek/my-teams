import * as React from 'react';
import styles from "./CSS/MyTeamsHome.module.scss";
import { escape } from '@microsoft/sp-lodash-subset';
import { ITeam, IChannel, ITeamImage, ITeamFinal } from "../../../shared/interfaces";
import { IMyTeamsHomeProps, IMyTeamsHomeState } from '.';
import DataTable, { createTheme } from 'react-data-table-component';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { Constants } from "../components/constants";
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
require('office-ui-fabric-react/dist/css/fabric.min.css');

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 116, border: 'none' },
};

const options: IDropdownOption[] = [
  { key: 'Header', text: 'Order by', itemType: DropdownMenuItemType.Header },
  { key: 'A-Z', text: 'A-Z' },
  { key: 'Z-A', text: 'Z-A' },
  { key: 'Latest', text: 'Latest' },
  { key: 'Oldest', text: 'Oldest' },
];


export default class MyTeamsHome extends React.Component<IMyTeamsHomeProps, IMyTeamsHomeState, {}> {
  private _myTeams: ITeam[] = null;
  private _filteredTeams: ITeam[] = null;
  private configurations: any;
  private constants = new Constants();

  constructor(props: IMyTeamsHomeProps) {
    super(props);

    this.state = {
      items: null,
      SortedBy: this.constants.SortedBy.NameAsc
    };
  }

  public async componentDidMount() {
    this.configurations = JSON.parse(sessionStorage.getItem(this.constants.CacheKey.TeamsData));
    if (!this.configurations) {
      await this._load().then(x => {
        sessionStorage.setItem(this.constants.CacheKey.TeamsData, JSON.stringify(this._myTeams));
        this.configurations = JSON.parse(sessionStorage.getItem(this.constants.CacheKey.TeamsData));
        this.setTeamsData(this._filteredTeams);
      })

    }
    else {
      this.setTeamsData(this._filteredTeams);
    }
  }

  public async componentDidUpdate(prevProps: IMyTeamsHomeProps) {
    if (this.props.openInClientApp !== prevProps.openInClientApp) {
      this.configurations = JSON.parse(sessionStorage.getItem(this.constants.CacheKey.TeamsData));
      if (!this.configurations) {
        await this._load().then(x => {
          sessionStorage.setItem(this.constants.CacheKey.TeamsData, JSON.stringify(this._myTeams));
          this.configurations = JSON.parse(sessionStorage.getItem(this.constants.CacheKey.TeamsData));
          this.setTeamsData(this._filteredTeams);
        })

      }
      else {
        console.log(this.configurations);
        this.setTeamsData(this._filteredTeams);
      }
    }

  }

  private _load = async () => {
    // get teams

    await this._getTeams().then(async (value) => {
      this._myTeams = value.filter(team => team.displayName != "Company Administrator").filter(team => team.resourceProvisioningOptions.length > 0);
      console.log(this._myTeams);
      for (let i = 0; i < this._myTeams.length; i++) {
        await this._getTeamChannels(this._myTeams[i].id).then(async channels => {
          this._myTeams[i].channelsData = channels;
        })
      }
    });
  }


  private setTeamsData = async (parameter) => {
    let teamsData: ITeam[] = parameter ? parameter : this.configurations;
    if (!parameter)
      for (let i = 0; i < teamsData.length; i++) {
        await this._getImagePath(teamsData[i].id).then(imagePath => {
          teamsData[i].imagePath = imagePath;
        });
      }

    this.setState({
      items: this.sortData(teamsData, this.state.SortedBy)
    });
  }


  private sortData = (data: ITeam[], parameter: string) => {
    switch (parameter) {

      case this.constants.SortedBy.NameAsc:
        data.sort((a, b) => a.displayName > b.displayName ? 1 : -1);
        break;

      case this.constants.SortedBy.NameDesc:
        data.sort((a, b) => a.displayName < b.displayName ? 1 : -1);
        break;

      case this.constants.SortedBy.RecentlyAddedAsc:
        this.sortBy(data, {
          prop: "createdDateTime",
          desc: false,
          parser: function (item) {
            return new Date(item);
          }
        });
        break;
      case this.constants.SortedBy.RecentlyAddedDesc:
        this.sortBy(data, {
          prop: "createdDateTime",
          desc: true,
          parser: function (item) {
            return new Date(item);
          }
        });
        break;
        break;
    }
    return data;
  }

  public sortByDate(arr) {
    arr.sort(function (a, b) {
      return Number(new Date(a.createdDateTime)) - Number(new Date(b.createdDateTime));
    });

    return arr;
  }

  public sortBy = (function () {
    var toString = Object.prototype.toString,
      // default parser function
      parse = function (x) { return x; },
      // gets the item to be sorted
      getItem = function (x) {
        var isObject = x != null && typeof x === "object";
        var isProp = isObject && this.prop in x;
        return this.parser(isProp ? x[this.prop] : x);
      };

    /**
     * Sorts an array of elements.
     *
     * @param {Array} array: the collection to sort
     * @param {Object} cfg: the configuration options
     * @property {String}   cfg.prop: property name (if it is an Array of objects)
     * @property {Boolean}  cfg.desc: determines whether the sort is descending
     * @property {Function} cfg.parser: function to parse the items to expected type
     * @return {Array}
     */
    return function sortby(array, cfg) {
      if (!(array instanceof Array && array.length)) return [];
      if (Object.prototype.toString.call(cfg) !== "[object Object]") cfg = {};
      if (typeof cfg.parser !== "function") cfg.parser = parse;
      cfg.desc = !!cfg.desc ? -1 : 1;
      return array.sort(function (a, b) {
        a = getItem.call(cfg, a);
        b = getItem.call(cfg, b);
        return cfg.desc * (a < b ? -1 : +(a > b));
      });
    };

  }());

  public render(): React.ReactElement<IMyTeamsHomeProps> {
    const CustomTitle = ({ row }) => (
      <div className={styles.row}>
        <div className={"ms-Grid-col ms-xl2 ms-sm2"}>
          <img src={row.imagePath} className={styles.teamImage}></img>
        </div>
        <div className={"ms-Grid-col ms-xl10 ms-sm10"}>
          <div style={{ padding: '20px' }}>
            {}
            <div style={{ color: 'rgb(64,64,64)', overflow: 'hidden', textOverflow: 'ellipses', fontSize: '14px', fontFamily: 'Montserrat', fontWeight: "bolder" }}> {}
              {/* <a href="#" title='Click to open channel' onClick={this._openChannel.bind(this, row.id)}> */}
              <span>{row.displayName}</span>
              {/* </a> */}
            </div>
            <div style={{ color: 'darkgrey', overflow: 'hidden', textOverflow: 'ellipses' }}>
              {}
              {row.visibility}
            </div>
          </div>
        </div>
      </div>
    );

    const columns = [
      {
        name: 'Team Name',
        selector: 'displayName',
        sortable: false,
        maxWidth: '500px',
        padding: '5px',
        cell: row => <CustomTitle row={row} />,
      }
    ];

    const paginationOptions = { rowsPerPageText: 'Rows/page', rangeSeparatorText: 'of', selectAllRowsItem: true, selectAllRowsItemText: 'All' }

    const ExpanableComponent = ((result: any) => <div className={styles.channelSection}>{this.ExpandableChannels(result.data.channelsData)}
    </div>);


    return (
      <div className={styles.container} >
        <div className={styles.row}>
          <div className={"ms-Grid-col ms-xl ms-sm8"}>
            <SearchBox className={styles["ms-SearchBox"]} placeholder="Search" onChange={newValue => this.serachTeams(newValue)} />
          </div>
          <div className={"ms-Grid-col ms-xl4 ms-sm4 " + styles["ms-SortControl"]}>
            <Dropdown
              placeholder="Order by"
              options={options}
              styles={dropdownStyles}
              onChange={this.updateSortBy}
              defaultSelectedKey={this.constants.SortedBy.NameAsc}
              selectedKey={this.state.SortedBy.replace(' ', '')}
              dropdownWidth={176}
            />
          </div>
        </div>
        <div className={styles.row}>
          <div className={"ms-Grid-col ms-xl12 ms-sm12"}>
            {!this.state.items ? <div className={styles.progressMessage}>Fetching teams data...</div> :
              (this.state.items.length > 0 ?
                <DataTable
                  data={this.state.items}
                  columns={columns}
                  pagination
                  noTableHead
                  noHeader
                  highlightOnHover
                  pointerOnHover
                  paginationPerPage={5}
                  paginationRowsPerPageOptions={[5, 10, 15, 20, 25]}
                  paginationComponentOptions={paginationOptions}
                  expandableRows
                  expandableRowsComponent={<ExpanableComponent />}
                  expandOnRowClicked
                />
                : <div className={styles.progressMessage}>Teams data not found...</div>)
            }
          </div >
        </div >
      </div >
    );
  }

  private updateSortBy = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.setState({
      SortedBy: item.text,
      items: this.sortData(this.state.items, item.text)
    });

  }

  private ExpandableChannels(Channels): JSX.Element {
    if (Channels)
      return Channels.map((channel) => {
        return (
          <div className={styles.channelLink} onClick={this._openChannel.bind(this, channel)}>
            <span>{channel.displayName}</span>
            <span className={styles.membershipType}>
              {}
              {channel.membershipType == "standard" ? "Public" : "Private"}
            </span>
          </div>
        )
      });
  }


  private serachTeams(keyword: string) {
    if (keyword.length > 2) {
      this._filteredTeams = this.configurations.filter(team => team.displayName.toLowerCase().search(keyword.trim().toLowerCase()) != -1);
      this.setTeamsData(this._filteredTeams);
    }
    else {
      this.setState({
        items: this.configurations
      });
    }
  }



  private _openChannel = async (channel): Promise<void> => {
    let link = '#';
    if (channel) {
      if (this.props.openInClientApp) {
        link = channel.webUrl;
      } else {
        link = `https://teams.microsoft.com/_#/conversations/${channel.displayName}?threadId=${channel.id}&ctx=channel`;
      }
      window.open(link, '_blank');
    }
  }


  private _getTeams = async (): Promise<ITeam[]> => {
    let myTeams: ITeam[] = [];
    try {
      myTeams = await this.props.teamsService.GetTeams();
      console.log(myTeams);
    } catch (error) {
      console.log('Error getting teams', error);
    }
    return myTeams;
  }

  private _getImagePath = async (teamID) => {
    let imagePath: string = "";
    await this._getImage(teamID).then(Response => {
      imagePath = Response;
    });
    return imagePath;
  }


  private _getImage = async (teamId) => {
    let teamImagePath = null;
    try {
      teamImagePath = this.props.teamsService.GetTeamPhoto(teamId);
      console.log(teamImagePath);
    } catch (error) {
      console.log('Error getting team image', error);
    }
    return teamImagePath;
  }

  private _getTeamChannels = async (teamId): Promise<IChannel[]> => {
    let channels: IChannel[] = [];
    try {
      channels = await this.props.teamsService.GetTeamChannels(teamId);
      console.log(channels);
    } catch (error) {
      console.log('Error getting channels for team ' + teamId, error);
    }
    return channels;
  }
}
