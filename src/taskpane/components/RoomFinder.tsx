import * as React from 'react';
// import { DefaultButton } from 'office-ui-fabric-react';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import RoomList from './RoomList';
import axios from 'axios';

export interface AppProps {
}

export interface AppState { 
  startTime: Date;
  endTime: Date;
  showUnavailable: boolean;
  roomData: Array<any>; // is it acceptable for this to be generic or should we pull in IRoomButtonProps
}

export default class RoomFinder extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      startTime: null,
      endTime: null,
      showUnavailable: false,
      roomData: [],
    };
  }


  componentDidMount() {
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

  click = async () => {
    var item = Office.context.mailbox.item;
    Promise.all([this.makePromise(item.start), this.makePromise(item.end)])
      .then(function(values) {
        console.log(values);
      })
      .catch(function(error) {
        console.log(error);
      });
  };
  
  onToggleChange = ({}, checked: boolean) => {
    this.setState({showUnavailable: !checked});

    // todo RT: this doesn't really belong here, but this is a convenient buttonable place for now
    axios.get('http://localhost:2999/spaces/rooms/availability?start=2019-10-10T08%3A00%3A00&end=2019-10-10T09%3A00%3A00')
      .then(response => {
        this.setState({
          ...this.state,
          roomData: response.data
        });
      }); // todo catch for error handling    
  };

  render() {
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
