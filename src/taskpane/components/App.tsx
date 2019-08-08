import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import Header from './Header';
import HeroList, { HeroListItem } from './HeroList';
import Progress from './Progress';

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: []
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
    /**
    * Insert your Outlook code here
    */
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

    return (
      <div className='ms-welcome'>
        <Header logo='assets/logo-filled.png' title={this.props.title} message='Welcome' />
        <HeroList message='Discover what Ad Astra for Outlook can do for you!' items={this.state.listItems}>
          <p className='ms-font-l'>Modify the source files, then click <b>Run</b>.</p>
          <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}>Run</Button>
        </HeroList>
      </div>
    );
  }
}
