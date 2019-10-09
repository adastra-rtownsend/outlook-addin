import * as React from 'react';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { List } from 'office-ui-fabric-react/lib/List';

import { DetailedRoomButton } from './DetailedRoomButton';
import { IRoomInfo } from '../../taskpane/components/DetailedRoomButton';
//import { IExampleItem } from '../../utilities/exampleData';


export interface IRoomListProps {
  items: IRoomInfo[];
  showUnavailable: boolean;
}

export interface IRoomListState {
}

const evenItemHeight = 25;
const oddItemHeight = 50;
const numberOfItemsOnPage = 10;

export default class RoomList extends React.Component<IRoomListProps, IRoomListState> {
  constructor(props: IRoomListProps) {
    super(props);

    this.state = {
    };
  }

  public render() {
    let items = this.props.items;
    const { showUnavailable } = this.props;

    if (!showUnavailable) {
      items = items.filter(item => item.available === true);
    }

    return (
      <FocusZone direction={FocusZoneDirection.vertical}>
        <div className='scroll-container' data-is-scrollable={true}>
          <List items={items} getPageHeight={this._getPageHeight} onRenderCell={this._onRenderCell} />
        </div>
      </FocusZone>
    );
  }

  private _getPageHeight(idx: number): number {
    let h = 0;
    for (let i = idx; i < idx + numberOfItemsOnPage; ++i) {
      const isEvenRow = i % 2 === 0;

      h += isEvenRow ? evenItemHeight : oddItemHeight;
    }
    return h;
  }

  private _onRenderCell = (item: IRoomInfo): JSX.Element => {
    return (
      <div data-is-focusable={true}>
        <DetailedRoomButton roomInfo={item} />
      </div>
    );
  };
}
