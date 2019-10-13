import * as React from 'react';
import { PrimaryButton, ButtonType } from 'office-ui-fabric-react';
import Header from './Header';
import HeroList, { HeroListItem } from './HeroList';
import RoomFinder from './RoomFinder';
import { WELCOME_SCREEN_SETTTING } from '../../utilities/config';
import { DEMO_DATA_SETTING } from '../../utilities/config';
import { API_PATH_SETTING } from '../../utilities/config';
import { getDefaultSettings} from '../../utilities/config';

export interface AppProps {
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
  haveInitializedOfficeSettings: boolean;
  showIntro: boolean;
  useSampleData: boolean;
  apiBasePath: string;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      haveInitializedOfficeSettings: false,
      // can't set these accurately until isOfficeInitialized is true 
      showIntro: true, 
      useSampleData: false,
      apiBasePath: '',
    };
  }

  _initializeOfficeSettings() {

    var defaults = getDefaultSettings();

    var showWelcome = Office.context.roamingSettings.get(WELCOME_SCREEN_SETTTING);
    var useSampleData = Office.context.roamingSettings.get(DEMO_DATA_SETTING);
    var apiBasePath = Office.context.roamingSettings.get(API_PATH_SETTING);

    if (showWelcome === undefined) {
      showWelcome = defaults.showWelcomeScreen;
      console.log(`Welcome screen setting not set, initializing to ${showWelcome}`);
    }
    
    if (useSampleData === undefined) {
      useSampleData = defaults.useSampleData;
      console.log(`Sample data setting not set, initializing to ${useSampleData}`);
    }
    
    if (apiBasePath === undefined) {
      apiBasePath = defaults.apiBasePath;
      console.log(`API base path setting not set, initializing to ${apiBasePath}`);
    }

    if (showWelcome === 1) {
      this.setState({showIntro: false});
    } else if (showWelcome === 2) {
      this.setState({showIntro: true});
      Office.context.roamingSettings.set(WELCOME_SCREEN_SETTTING, 1); // set to 'never' so it won't show next time
      Office.context.roamingSettings.saveAsync();
    } else {
      this.setState({showIntro: true});
    }

    this.setState({useSampleData: useSampleData});
    this.setState({apiBasePath: apiBasePath});
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: 'Search',
          primaryText: 'Find available rooms in Ad Astra'
        },
        {
          icon: 'DateTime',
          primaryText: 'Reserve room(s) in Ad Astra'
        }
      ]
    });
  }

  componentDidUpdate() {
    if (this.props.isOfficeInitialized && !this.state.haveInitializedOfficeSettings) {
      this._initializeOfficeSettings();
      this.setState({haveInitializedOfficeSettings: true});
      this.forceUpdate();
    } 
  }

  click = async () => {
    this.setState({ showIntro: false });
  }

  renderIntro() {
    return (
      <div className='ms-welcome'>
        <Header logo='assets/logo-filled.png' title='' message='Welcome' />
        <HeroList message='Discover what Ad Astra for Outlook can do for you!' items={this.state.listItems}>
          <PrimaryButton className='ms-welcome__action' buttonType={ButtonType.hero} onClick={this.click} text="Get Started"/>
        </HeroList>
      </div>
    );
  }
  
  render() {

    const {
      isOfficeInitialized,
    } = this.props;

    if (!isOfficeInitialized || !this.state.haveInitializedOfficeSettings) {
      return (
        <div></div>
      );
    } 
    
    if (this.state.showIntro) {
      return this.renderIntro();
    } else {
      return (
        <RoomFinder 
          useSampleData={this.state.useSampleData}
          apiBasePath={this.state.apiBasePath}
        >
        </RoomFinder>
      );
    }
  }
}
