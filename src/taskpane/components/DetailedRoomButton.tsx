import * as React from 'react';
import { IPersonaSharedProps, Persona, PersonaSize, PersonaPresence, PersonaInitialsColor } from 'office-ui-fabric-react/lib/Persona';
import { ActionButton } from 'office-ui-fabric-react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';

export const DetailedRoomButton: React.FunctionComponent = () => {
  const examplePersona: IPersonaSharedProps = {
    showUnknownPersonaCoin: true,
    text: '1st Floor - Engineering Conference Room',
    secondaryText: 'Available',
    tertiaryText: 'Not Available',
    showSecondaryText: true
  };
  
  return (
    <ActionButton allowDisabledFocus onClick={() => _addLocation(examplePersona.text)} style={{ paddingLeft: '16px', paddingRight: '16px', 
                  paddingBottom: '9px', paddingTop: '9px', height: 'auto'}}>
      <Persona {...examplePersona} size={PersonaSize.size32} presence={PersonaPresence.none} 
          onRenderSecondaryText={_onRenderSecondaryText}
          onRenderInitials ={_onRenderInitials} />
    </ActionButton>
  );

  function _onRenderInitials(props: IPersonaProps): JSX.Element {
    return (
      <Icon iconName="Room"/>
    );
  };

  function _onRenderSecondaryText(props: IPersonaProps): JSX.Element {
    return (
      <div>
        <span className='availibity-text'>
          <Icon iconName="Clock" styles={{ color: '0xff0000', root: { marginRight: 5 } }} />
          {props.secondaryText}
        </span>
        <span>
          <Icon iconName="Contact" styles={{ root: { marginRight: 5 } }} />
          <span>34</span> 
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