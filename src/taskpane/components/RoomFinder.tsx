import * as React from 'react';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import RoomList from './RoomList';
import axios from 'axios';
import { Spinner, SpinnerSize, PrimaryButton, ButtonType } from 'office-ui-fabric-react';
import { Stack, IStackStyles } from 'office-ui-fabric-react/lib/Stack';
import * as moment from 'moment';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';
import SettingsDialog from './SettingsDialog';
import { number } from 'prop-types';

const stackStyles: IStackStyles = {
  root: {
    height: 250
  }
};

export interface IRoomFinderProps {
}

export interface IRoomFinderState { 
  isLoading: boolean;
  hasError: boolean;
  startTime: any;
  endTime: any;
  showUnavailable: boolean;
  roomData: Array<any>; // is it acceptable for this to be generic or should we pull in IRoomButtonProps
  settingsDialog?: React.RefObject<SettingsDialog>;
}

export default class RoomFinder extends React.Component<IRoomFinderProps, IRoomFinderState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      isLoading: false,
      hasError: false, 
      startTime: null,
      endTime: null,
      showUnavailable: false,
      roomData: [],
      settingsDialog: React.createRef(),
    };
  }

  onInterval() {
    this.refreshRoomInfo(false);
  }  
  
  componentDidMount() {
    setInterval(this.onInterval.bind(this), 2000);
    this.refreshRoomInfo(true);
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

  refreshRoomInfo = async (force) => {
    var that = this; 
    var item = Office.context.mailbox.item;

    Promise.all([that.makePromise(item.start), that.makePromise(item.end)])
      .then(function(values) {
        if (force || !moment(values[0]).isSame(that.state.startTime) || !moment(values[1]).isSame(that.state.endTime)) {
          that.setState({isLoading: true});
          that.setState({startTime: moment(values[0])});
          that.setState({endTime: moment(values[1])});
          var startTime = encodeURIComponent(moment(values[0]).format('YYYY-MM-DDTHH:mm:ss'));
          var endTime = encodeURIComponent(moment(values[1]).format('YYYY-MM-DDTHH:mm:ss'));
          var url = `http://localhost:2999/spaces/rooms/availability?start=${startTime}&end=${endTime}`;
          axios.get(url)
            .then(response => {
              this.getCapacity().then(capacity => {
                console.log(capacity)
                const roomData = []
                response.data.forEach((room) => {
                  roomData.push({
                    ...room,
                    capacity: capacity[room.id].capacity
                  });
                });
                that.setState({isLoading: false});
                that.setState({
                  ...that.state,
                  roomData
                });
              // });
            }); // todo neeed error handling, shouldn't jsut assume API succeeds
          }
        })
      .catch(function(error) {
        console.log(error);
        that.setState({isLoading: false});
      });
  }
  
  getCapacity() {
    let capacity = {};

    // Async way
    // try {
    //   const response = await axios.get('http://localhost:2999/facilities/roomlist?filtertype=equals_%2F_in')
    //   response.data.forEach((room) => {
    //     capacity.push(room[9]);
    //   });

    //   return capacity;
    // } catch (err) {
    //   return err
    // }

    return new Promise((resolve, reject) => {
      axios.get('http://localhost:2999/facilities/roomlist?filtertype=equals_%2F_in').then(response => {
        response.data.forEach((room) => {
          capacity[room["roomId"]] = {
            capacity: room["maxOccupancy"]
          }
        });
        resolve(capacity);
      }).catch(err => {
        reject(err);
      });
    });
  };

  onToggleChange = ({}, checked: boolean) => {
    this.setState({showUnavailable: !checked});
  };

  dismissError = () => {
    this.setState({hasError: false});
  };

  onBookRoom = () => {
    let roomData = Office.context.roamingSettings.get('selectedRoom');
    if (roomData && roomData.text) {
      var that = this; 
      var startTime = encodeURIComponent(moment(this.state.startTime).format('YYYY-MM-DDTHH:mm:ss'));
      var endTime = encodeURIComponent(moment(this.state.endTime).format('YYYY-MM-DDTHH:mm:ss'));
      var roomId = roomData.roomId;

      var url = `http://localhost:2999/spaces/rooms/${roomId}/reservation/?start=${startTime}&end=${endTime}`;
      axios.post(url).then(() => {
        that.setState({hasError: false});
        Office.context.mailbox.item.location.setAsync(roomData.text, function (asyncResult) {
          if (asyncResult.status == Office.AsyncResultStatus.Failed) {
              console.log("Error written location in outlook : " + asyncResult.error.message);
          } else {
              console.log("Location written in outlook");
          }
        });
      }).catch(error => {
        this.setState({hasError: true});
        console.log(error);
      }); // todo neeed error handling, shouldn't jsut assume API succeeds                
    }
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
                      borderColor: 'rgba(237, 235, 233, 1)', borderBottomStyle: 'solid'}}
                      onContextMenu={(e) => { 
                        this.state.settingsDialog.current.showDialog();
                        e.preventDefault();
                      }}                
        >        
          <SettingsDialog ref={this.state.settingsDialog} />       
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
        { !this.state.hasError && 
          <PrimaryButton className='book-room-button' buttonType={ButtonType.hero} onClick={this.onBookRoom} text="Book Room"/>
        }
        { this.state.hasError && 
          <MessageBar className='error-message-bar' messageBarType={MessageBarType.error} 
            onDismiss={this.dismissError} isMultiline={false} dismissButtonAriaLabel="Close">
              Failed to book room in Astra Schedule
          </MessageBar>
        }    
      </div>
);
  }
}
