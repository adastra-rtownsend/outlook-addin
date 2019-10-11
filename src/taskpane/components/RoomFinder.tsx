import * as React from 'react';
// import { DefaultButton } from 'office-ui-fabric-react';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import RoomList from './RoomList';
import axios from 'axios';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { Stack, IStackStyles } from 'office-ui-fabric-react/lib/Stack';
import * as moment from 'moment';

const stackStyles: IStackStyles = {
  root: {
    height: 250
  }
};

export interface IRoomFinderProps {
}

export interface IRoomFinderState { 
  isLoading: boolean;
  startTime: Date;
  endTime: Date;
  showUnavailable: boolean;
  roomData: Array<any>; // is it acceptable for this to be generic or should we pull in IRoomButtonProps
}

export default class RoomFinder extends React.Component<IRoomFinderProps, IRoomFinderState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      isLoading: false,
      startTime: null,
      endTime: null,
      showUnavailable: false,
      roomData: [],
    };
  }


  componentDidMount() {
    this.refreshRoomInfo();
  }

  makePromise = function (itemField) {
    return new Promise(function(resolve, reject) {
      itemField.getAsync(function (asyncResult) {
        if (asyncResult.status.toString === "failed") {
          reject(asyncResult.error.message);
        }
        else {
          resolve(asyncResult.value);
        }
      });
    });
  }

  refreshRoomInfo = () => {
    var that = this; 
    that.setState({isLoading: true});
    var item = Office.context.mailbox.item;
    Promise.all([that.makePromise(item.start), that.makePromise(item.end)])
      .then(function(values) {
        var startTime = encodeURIComponent(moment(values[0]).format('YYYY-MM-DDTHH:mm:ss'));
        var endTime = encodeURIComponent(moment(values[1]).format('YYYY-MM-DDTHH:mm:ss'));
        var url = `http://localhost:2999/spaces/rooms/availability?start=${startTime}&end=${endTime}`;
        console.log(url);
        axios.get(url)
          .then(response => {
            that.setState({
              ...that.state,
              roomData: response.data
            });
            that.setState({isLoading: false});
          }); // todo neeed error handling, shouldn't jsut assume API succeeds
      })
      .catch(function(error) {
        console.log(error);
        that.setState({isLoading: false});
      });
  }

  click = async () => {
  };
  
  onToggleChange = ({}, checked: boolean) => {
    this.setState({showUnavailable: !checked});

    // todo RT: this doesn't really belong here, but for now this is a convenient place to force a refresh
    this.refreshRoomInfo();
  };

  render() {
    if (this.state.isLoading) {
      return (
        <Stack grow>
          <Stack verticalAlign="center" styles={stackStyles}>
          <Spinner size={SpinnerSize.large} label="Loading room data" ariaLive="assertive" labelPosition="right" />
          </Stack>
        </Stack>
      )
    }

    return (
      <div>
        <div style={{ paddingLeft: '16px', paddingRight: '16px', paddingBottom: '10px', borderBottomWidth: '1px',
                       borderColor: 'rgba(237, 235, 233, 1)', borderBottomStyle: 'solid'}}>        
          <div className="ms-SearchBoxExample" style={{borderColor: 'rgba(237, 235, 233, 1)'}}>
            <SearchBox
              placeholder="Search by Ad Astra room name"
              onSearch={newValue => console.log('value is ' + newValue)}
              onFocus={() => console.log('onFocus called')}
              onBlur={() => console.log('onBlur called')}
              onChange={() => console.log('onChange called')}
            />
          </div>
          <div style={{ marginTop: '13px', marginBottom: '5px' }} > 
              <Toggle
                defaultChecked={!this.state.showUnavailable}
                label="Only available rooms"
                inlineLabel={true}
                onFocus={() => console.log('onFocus called')}
                onBlur={() => console.log('onBlur called')}
                onChange={this.onToggleChange}
              />
          </div>
        </div>
        <RoomList items={this.state.roomData} showUnavailable={this.state.showUnavailable} />

        {/* <DefaultButton className='ms-welcome__action'  onClick={this.click} text="Refresh"/>
        <div>Here is what I pulled of invite: {JSON.stringify(this.state)} </div> */}
      </div>
);
  }
}
