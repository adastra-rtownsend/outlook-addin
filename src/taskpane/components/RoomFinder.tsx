import * as React from 'react';
import { Button } from 'office-ui-fabric-react';

export interface AppProps {
}

export interface AppState { 
  startTime: Date;
  endTime: Date;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      startTime: null,
      endTime: null,
    };
  }


  componentDidMount() {
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

  click = async () => {

    var item = Office.context.mailbox.item;
    Promise.all([this.makePromise(item.start), this.makePromise(item.end)])
      .then(function(values) {
        console.log(values);
      })
      .catch(function(error) {
        console.log(error);
      });
  }
  
  render() {
    return (
      <div>
        <Button className='ms-welcome__action'  onClick={this.click}>Refresh</Button>
        <div>Here is what I pulled of invite: {JSON.stringify(this.state)} </div>
      </div>
);
  }
}
