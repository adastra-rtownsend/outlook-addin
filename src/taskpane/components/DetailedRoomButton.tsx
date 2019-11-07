import * as React from 'react';
import { IPersonaSharedProps, Persona, PersonaSize, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import { CompoundButton } from 'office-ui-fabric-react';
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
  onClickFn: Function;
}

export interface IRoomButtonState {
  selected: boolean;
}

export class DetailedRoomButton extends React.Component<IRoomButtonProps, IRoomButtonState> {
  constructor(props: IRoomButtonProps) {
    super(props);

    this.state = {
      selected: false,
    }
  }

  roomPersona: IRoomInfoProps = {
    showUnknownPersonaCoin: true,
    text: this.props.roomInfo.roomBuildingAndNumber,
    showSecondaryText: true,
    roomId: this.props.roomInfo.roomId,
    available: (true === this.props.roomInfo.available),
    capacity: this.props.roomInfo.capacity,
  }

  public render() {
    return (
      <CompoundButton checked={this.state.selected} allowDisabledFocus onClick={() => this.props.onClickFn(this) } style={{
                    paddingBottom: '9px', paddingTop: '9px', height: 'auto', width: '100%',
                    borderStyle: 'none', alignItems: 'start', textAlign: 'left', maxWidth: '500px'
                    }}>
        <Persona {...this.roomPersona} size={PersonaSize.size32} presence={PersonaPresence.none}
            onRenderSecondaryText={this._onRenderSecondaryText}
            onRenderInitials ={this._onRenderInitials}
            style={{

            }}
            />
      </CompoundButton>
    );
  }

  _onRenderInitials(): JSX.Element {
    return (
      <Icon iconName="Room"/>
    );
  };

  _onRenderSecondaryText(props: IRoomInfoProps): JSX.Element {

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

  _selectRoom(roomData): void {
    Office.context.roamingSettings.set(SELECTED_ROOM_SETTING, roomData);
  };
}
