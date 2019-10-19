import * as React from 'react';
import { IPersonaSharedProps, Persona, PersonaSize, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import { ActionButton } from 'office-ui-fabric-react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { SELECTED_ROOM_SETTING } from '../../utilities/config';

export interface IRoomInfoProps extends IPersonaSharedProps {
  roomId: string;
  available: boolean; 
  capacity?: number;
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
  };
  
  return (
    <ActionButton allowDisabledFocus onClick={() => _selectRoom(roomPersona)} style={{ paddingLeft: '16px', paddingRight: '16px', 
                  paddingBottom: '9px', paddingTop: '9px', height: 'auto'}}>
      <Persona {...roomPersona} size={PersonaSize.size32} presence={PersonaPresence.none} 
          onRenderSecondaryText={_onRenderSecondaryText}
          onRenderInitials ={_onRenderInitials} />
    </ActionButton>
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
    Office.context.roamingSettings.set(SELECTED_ROOM_SETTING, roomData);
  };
};

