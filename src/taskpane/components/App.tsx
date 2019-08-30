import * as React from 'react';
import { PrimaryButton, ButtonType } from 'office-ui-fabric-react';
import Header from './Header';
import HeroList, { HeroListItem } from './HeroList';
import Progress from './Progress';
import RoomFinder from './RoomFinder';

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
  showIntro: boolean;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      showIntro: true,
    };
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

  click = async () => {
    this.setState({ showIntro: false });
  }

  renderIntro() {
    return (
      <div className='ms-welcome'>
        <Header logo='assets/logo-filled.png' title={this.props.title} message='Welcome' />
        <HeroList message='Discover what Ad Astra for Outlook can do for you!' items={this.state.listItems}>
          <PrimaryButton className='ms-welcome__action' buttonType={ButtonType.hero} onClick={this.click} text="Get Started">Get Started</PrimaryButton>
        </HeroList>
      </div>
    );
  }
  
  render() {

    const {
      title,
      isOfficeInitialized,
    } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo='assets/logo-filled.png'
          message='Please sideload your addin to see app body.'
        />
      );
    }

    if (this.state.showIntro) {
      return this.renderIntro();
    } else {
      return (
        <RoomFinder></RoomFinder>
      );
    }
  }
}
