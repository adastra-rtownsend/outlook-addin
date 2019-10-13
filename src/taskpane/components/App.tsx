import * as React from 'react';
import { PrimaryButton, ButtonType } from 'office-ui-fabric-react';
import Header from './Header';
import HeroList, { HeroListItem } from './HeroList';
import RoomFinder from './RoomFinder';

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
    var showWelcome = Office.context.roamingSettings.get('adastra.showWelcomeScreen');
    if (showWelcome === 1) {
      this.setState({showIntro: false});
    } else if (!showWelcome || showWelcome === 2) { // also handle thie case where setting doesn't exist yet
      this.setState({showIntro: true});
      Office.context.roamingSettings.set('adastra.showWelcomeScreen', 1); // set to 'never' so it won't show nexxt time
      Office.context.roamingSettings.saveAsync();
    } else {
      this.setState({showIntro: true});
    }

    this.setState({useSampleData: false !== Office.context.roamingSettings.get('adastra.useSampleData')});
    this.setState({apiBasePath: Office.context.roamingSettings.get('adastra.apiBasePath')});
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
