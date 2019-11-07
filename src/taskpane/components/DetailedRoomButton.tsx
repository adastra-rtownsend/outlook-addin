import * as React from 'react';
import { IPersonaSharedProps, Persona, PersonaSize, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import { CompoundButton } from 'office-ui-fabric-react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { SELECTED_ROOM_SETTING } from '../../utilities/config';

export interface IRoomInfoProps extends IPersonaSharedProps {
  roomId: string;
  available: boolean;
  capacity?: number;
  selected: boolean;
}

// note: this should match to server definition
export interface ISourceRoomInfo {
  roomId: string;
  roomBuildingAndNumber: string;
  whyIsRoomIdHereTwice: string;
  available: boolean;
  capacity?: number;
}

export interface IRoomButtonProps {
  roomInfo: ISourceRoomInfo;
 }

export const DetailedRoomButton: React.SFC<IRoomButtonProps> = (props) => {
  const roomPersona: IRoomInfoProps = {
    showUnknownPersonaCoin: true,
    text: props.roomInfo.roomBuildingAndNumber,
    showSecondaryText: true,
    roomId: props.roomInfo.roomId,
    available: (true === props.roomInfo.available),
    capacity: props.roomInfo.capacity,
    selected: false,
  };

  return (
    // checked={roomPersona.selected}
    <CompoundButton checked={roomPersona.selected} allowDisabledFocus onClick={() => _selectRoom(roomPersona) } style={{
                  paddingBottom: '9px', paddingTop: '9px', height: 'auto', width: '100%',
                  borderStyle: 'none', alignItems: 'start', textAlign: 'left', maxWidth: '500px'
                  }}>
      <Persona {...roomPersona} size={PersonaSize.size32} presence={PersonaPresence.none}
          onRenderSecondaryText={_onRenderSecondaryText}
          onRenderInitials ={_onRenderInitials}
          style={{

          }}
          />
    </CompoundButton>
  );

  function _onRenderInitials(): JSX.Element {
    return (
      <Icon iconName="Room"/>
    );
  };

  function _onRenderSecondaryText(props: IRoomInfoProps): JSX.Element {

    let clockIcon = 'Clock';
    let text = 'Available';
    let style = 'available-text';

    if (props.available === false) {
      clockIcon = 'CircleStopSolid';
      text = 'Unavailable';
      style = 'unavailable-text';
    }

    return (
      <div>
        <span className={style}>
          <Icon iconName={clockIcon} styles={{ root: { marginRight: 5 } }} />
          {text}
        </span>
        { props.capacity &&
          <span>
            <Icon iconName="Contact" styles={{ root: { marginRight: 5 } }} />
            <span>{props.capacity}</span>
          </span>
        }
      </div>
    );
  };

  function _selectRoom(roomData): void {
    console.log(JSON.stringify(roomData, null, 2));
    const was = roomData.selected;
    roomData.selected = was ? false : true;
    //console.log(`inside _selectRoom. roomData.selected was=${was} now=${!roomData.selected}`);
    Office.context.roamingSettings.set(SELECTED_ROOM_SETTING, roomData);
  };
};

