import * as React from 'react';

export interface AppProps {
}

export interface AppState { 
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
  }

  componentDidMount() {
  }

  render() {
    return (
      <div>Room finder will go here</div>
    );
  }
}
