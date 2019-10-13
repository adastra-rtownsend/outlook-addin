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
  officeSettingsInitializationState: number; // quick hack - 0: unstarted 1: inprogress 2: done
  showIntro: boolean;
  useSampleData: boolean;
  apiBasePath: string;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      officeSettingsInitializationState: 0,
      // can't set these accurately until isOfficeInitialized is true 
      showIntro: true, 
      useSampleData: false,
      apiBasePath: '',
    };
  }

  _initializeOfficeSettings() {

    var defaults = getDefaultSettings();

    var showWelcome = Office.context.roamingSettings.get(WELCOME_SCREEN_SETTTING);
    var useDemoData = Office.context.roamingSettings.get(DEMO_DATA_SETTING);
    var apiPath = Office.context.roamingSettings.get(API_PATH_SETTING);

    if (showWelcome === undefined) {
      showWelcome = defaults.showWelcomeScreen;
      console.log(`Welcome screen setting not set, initializing to ${showWelcome}`);
      this.setState({showIntro: showWelcome});
      Office.context.roamingSettings.set(WELCOME_SCREEN_SETTTING, showWelcome);
    }
    
    if (useDemoData === undefined) {
      useDemoData = defaults.useSampleData;
      console.log(`Sample data setting not set, initializing to ${useDemoData}`);
      this.setState({useSampleData: useDemoData});
      Office.context.roamingSettings.set(DEMO_DATA_SETTING, useDemoData);
    }
    
    if (apiPath === undefined) {
      apiPath = defaults.apiBasePath;
      console.log(`API base path setting not set, initializing to ${apiPath}`);
      this.setState({apiBasePath: apiPath});
      Office.context.roamingSettings.set(API_PATH_SETTING, apiPath);
    }

    if (showWelcome === 1) {
      this.setState({showIntro: false});
    } else if (showWelcome === 2) {
      this.setState({showIntro: true});
      Office.context.roamingSettings.set(WELCOME_SCREEN_SETTTING, 1); // set to 'never' so it won't show next time
    } else {
      this.setState({showIntro: true});
    }

    // some changes may have occured, so sync the settings
    var that = this;
    Office.context.roamingSettings.saveAsync(() => {
      that.setState({officeSettingsInitializationState: 2});
    });
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
    if (this.props.isOfficeInitialized && this.state.officeSettingsInitializationState === 0) {
      this.setState({officeSettingsInitializationState: 1});
      this._initializeOfficeSettings();
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

    if (!isOfficeInitialized || this.state.officeSettingsInitializationState < 2) {
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
