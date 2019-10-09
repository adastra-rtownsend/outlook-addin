import * as React from 'react';
import { IPersonaSharedProps, Persona, PersonaSize, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import { ActionButton } from 'office-ui-fabric-react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
// import { IExampleItem } from '../../utilities/exampleData';

export interface IRoomInfoProps extends IPersonaSharedProps {
  available: boolean, 
  capacity: number
}

export interface IRoomInfo {
  key: string;
  roomName: string;
  roomBuilding: string;
  roomNumber: string;
  available: boolean;
  capacity: number;
};

interface IRoomButtonProps {
  roomInfo: IRoomInfo,
 }

export const DetailedRoomButton: React.SFC<IRoomButtonProps> = (props) => {
  const roomPersona: IRoomInfoProps = {
    showUnknownPersonaCoin: true,
    text: `${props.roomInfo.roomBuilding} - ${props.roomInfo.roomNumber}`,
    showSecondaryText: true,
    available: (true === props.roomInfo.available),
    capacity: props.roomInfo.capacity,
  };
  
  return (
    <ActionButton allowDisabledFocus onClick={() => _addLocation(roomPersona.text)} style={{ paddingLeft: '16px', paddingRight: '16px', 
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
        <span>
          <Icon iconName="Contact" styles={{ root: { marginRight: 5 } }} />
          <span>{props.capacity}</span> 
        </span>
      </div>
    );
  }; 
};

function _addLocation(roomName): void {  
  Office.context.mailbox.item.location.setAsync(roomName, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        console.log("Error written location in outlook : " + asyncResult.error.message);
    } else {
        console.log("Location written in outlook");
    }
  });
};