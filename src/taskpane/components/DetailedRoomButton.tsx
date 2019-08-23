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
    <ActionButton allowDisabledFocus style={{ paddingLeft: '16px', paddingRight: '16px', 
                  paddingBottom: '9px', paddingTop: '9px', height: 'auto'}}>
      <Persona {...examplePersona} size={PersonaSize.size32} presence={PersonaPresence.none} 
          onRenderInitials ={() => { return (<Icon iconName="Room"/>) } />
    </ActionButton>
  );
};
