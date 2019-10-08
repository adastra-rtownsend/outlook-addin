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
// const _cachedItems = []

export interface AppProps {
}

export interface AppState {
  startTime: Date;
  endTime: Date;
  showUnavailable: boolean;
  rooms: Array<any>;
}

export default class RoomFinder extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      startTime: null,
      endTime: null,
      showUnavailable: false,
      rooms: []
    };
  }

  componentDidMount() {
    this.getAvailableRooms();
  }

  getAvailableRooms() {
    let items = []
    axios.get('https://www.aaiscloud.com/AustinCC/~api/search/room?_dc=1570564904737&start=0&limit=500&_s=1&fields=RowNumber%2CId%2CRoomName%2CRoomDescription%2CRoomNumber%2CRoomTypeName%2CBuildingCode%2CBuildingName%2CCampusName%2CCapacity%2CBuildingRoomNumberRoomName%2CEffectiveDateId%2CCanEdit%2CCanDelete&sortOrder=%2BBuildingRoomNumberRoomName&page=1&sort=%5B%7B%22property%22%3A%22BuildingRoomNumberRoomName%22%2C%22direction%22%3A%22ASC%22%7D%5D').then(response => {
      response.data.data.forEach((d: any[]) => {
        items.push({
          key: d[0],
          id: d[1],
          roomName: d[2],
          roomNumber: d[4],
          roomBuilding: d[7],
          available: true,
          capacity: d[9]
        })
      });

      this.setState({
        ...this.state, 
        rooms: items
      }) 
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
        <RoomList items={this.state.rooms} showUnavailable={this.state.showUnavailable} />

        {/* <DefaultButton className='ms-welcome__action'  onClick={this.click} text="Refresh"/>
        <div>Here is what I pulled of invite: {JSON.stringify(this.state)} </div> */}
      </div>
);
  }
}
