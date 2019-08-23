import * as React from 'react';

export interface DetailedRoomButtonProps {
}

export interface DetailedRoomButtonState { 
}

export default class DetailedRoomButton extends React.Component<DetailedRoomButtonProps, DetailedRoomButtonState> {
  constructor(props, context) {
    super(props, context);
  }

  componentDidMount() {
  }
  
  render() {
    return (
      <div>
        This is a room
      </div>
    );
  }
}
