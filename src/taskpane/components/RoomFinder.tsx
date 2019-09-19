import * as React from 'react';
// import { DefaultButton } from 'office-ui-fabric-react';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';

import RoomList from './RoomList';
// import { createListItems } from '../../utilities/exampleData';

import axios from 'axios';
// import { response } from 'express';

// for now we are using fake data items in the scrolling room grid
// const _cachedItems = createListItems(5000);
const _cachedItems = []

export interface AppProps {
}

export interface AppState { 
  startTime: Date;
  endTime: Date;
  showUnavailable: boolean;
}

export default class RoomFinder extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      startTime: null,
      endTime: null,
      showUnavailable: false,
    };
  }

  componentDidMount() {
    this.getAvailableRooms();
  }

  getAvailableRooms() {
    axios.get('http://qeapp/SG86044Merced/~api/query/room?&fields=Id%2CName%2CroomNumber%2CRoomType%2EName%2CBuilding%2EName%2CBuilding%2EBuildingCode%2CMaxOccupancy%2CIsActive&allowUnlimitedResults=false&sort=%2BBuilding%2EName,Name&page=1&start=0&limit=200').then(response => {
      response.data.data.forEach((d: any[]) => {
        _cachedItems.push({
          key: d[0],
          roomName: d[1],
          roomNumber: d[2],
          roomBuilding: d[4],
          available: true,
          capacity: 100
        })
      });
    })
  }

  postReservation() {
    // Need to know when the event creation is saved in outlook
    console.log("Writing reservation back to Ad Astra")
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
        <RoomList items={_cachedItems} showUnavailable={this.state.showUnavailable} />

        {/* <DefaultButton className='ms-welcome__action'  onClick={this.click} text="Refresh"/>
        <div>Here is what I pulled of invite: {JSON.stringify(this.state)} </div> */}
      </div>
);
  }
}
