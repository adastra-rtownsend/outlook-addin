import * as React from 'react';

export interface DetailedRoomButtonProps {
  roomName: string;
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

    const {
      roomName,
    } = this.props;

    return (
      <div>
        This is a room: {roomName}
      </div>
    );
  }
  
}
