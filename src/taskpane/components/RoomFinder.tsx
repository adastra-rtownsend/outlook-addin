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
    this.getActivities(/* start and end time */);
    this.getAvailableRooms();
  }

  getActivities() {
    let mapRoomtoActivity = {};
    axios.get('https://qeapp/SG86044Merced/~api/calendar/activityList?fields=ActivityId%2CActivityName%2CStartDate%2CActivityTypeCode%2CCampusName%2CBuildingCode%2CRoomNumber%2CLocationName%2CStartDateTime%2CEndDateTime%2CInstructorName%3Astrjoin2(%22%20%22%2C%20%22%20%22%2C%20%22%20%22)%2CDays%3Astrjoin2(%22%20%22%2C%20%22%20%22%2C%20%22%20%22)%2CCanView%3Astrjoin2(%22%20%22%2C%20%22%20%22%2C%20%22%20%22)%2CSectionId%2CEventId%2CEventImage%3Astrjoin2(%22%20%22%2C%20%22%20%22%2C%20%22%20%22)%2CParentActivityId%2CParentActivityName%2CEventMeetingByActivityId%2EEvent%2EEventType%2EName%2CEventMeetingByActivityId%2EEventMeetingType%2EName%2CSectionMeetInstanceByActivityId%2ESectionMeeting%2EMeetingType%2EName%2CLocation%2ERoomId&filter=(((StartDateTime<%3D"2019-10-07T15%3A00%3A00")%26%26(EndDateTime>%3D"2019-10-07T14%3A00%3A00"))%7C%7C((StartDateTime>%3D"2019-10-07T14%3A00%3A00")%26%26(StartDateTime<%3D"2019-10-07T15%3A00%3A00")))&allowUnlimitedResults=false&sort=StartDateTime&page=1&start=0').then(response => {
      response.data.data.forEach((d: any[]) => {
        mapRoomtoActivity[d[21]] = true;
        });
      });
      return mapRoomtoActivity
    }

  getAvailableRooms() {
    let items = []
    const activities = this.getActivities();
    axios.get('https://qeapp/SG86044Merced/~api/search/room?_dc=1570564904737&start=0&limit=5000&_s=1&fields=RowNumber%2CId%2CRoomName%2CRoomDescription%2CRoomNumber%2CRoomTypeName%2CBuildingCode%2CBuildingName%2CCampusName%2CCapacity%2CBuildingRoomNumberRoomName%2CEffectiveDateId%2CCanEdit%2CCanDelete&sortOrder=%2BBuildingRoomNumberRoomName&sort=%5B%7B%22property%22%3A%22BuildingRoomNumberRoomName%22%2C%22direction%22%3A%22ASC%22%7D%5D').then(response => {
      response.data.data.forEach((d: any[]) => {
        items.push({
          key: d[0],
          id: d[1],
          roomName: d[2],
          roomNumber: d[4],
          roomBuilding: d[7],
          available: typeof activities[d[1]] === 'undefined',
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
