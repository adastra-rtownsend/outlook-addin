import * as React from 'react';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { List } from 'office-ui-fabric-react/lib/List';
import { DetailedRoomButton, ISourceRoomInfo } from './DetailedRoomButton';
import { ITheme, mergeStyleSets, getTheme, getFocusStyle } from 'office-ui-fabric-react/lib/Styling';

export interface IRoomListProps {
  items: ISourceRoomInfo[];
  showUnavailable: boolean;
}

export interface IRoomListState {
  selectedItem: DetailedRoomButton;
}

const evenItemHeight = 25;
const oddItemHeight = 50;
const numberOfItemsOnPage = 10;

const theme: ITheme = getTheme();
const { palette } = theme;

interface IListBasicExampleClassObject {
  itemCell: string;
}

const classNames: IListBasicExampleClassObject = mergeStyleSets({
  itemCell: [
    getFocusStyle(theme, { inset: -1 }),
    {
      selectors: {
        '&:hover': { background: palette.neutralLight }
      }
    }
  ],
  emptyList: [

  ]
});

export default class RoomList extends React.Component<IRoomListProps, IRoomListState> {
  constructor(props: IRoomListProps) {
    super(props);

    this.state = {
      selectedItem: null, // no initial selected item
    };
  }

  public render() {
    let items = this.props.items;
    const { showUnavailable } = this.props;

    if (!showUnavailable) {
      items = items.filter(item => item.available === true);
    }

    if (items.length > 0) {
      return (
        <FocusZone direction={FocusZoneDirection.vertical}>
          <div className='scroll-container' data-is-scrollable={true}>
            <List items={items} getPageHeight={this._getPageHeight} onRenderCell={this._onRenderCell} />
          </div>
        </FocusZone>
      );
    } else {
      return (
        <div className='empty-rooms-list'
             style={{ marginTop: '13px', marginBottom: '5px', marginLeft: '10px'}}
        >
          <h3>We couldn't find any rooms</h3>
        </div>
      );
    }
  }

  private _getPageHeight(idx: number): number {
    let h = 0;
    for (let i = idx; i < idx + numberOfItemsOnPage; ++i) {
      const isEvenRow = i % 2 === 0;

      h += isEvenRow ? evenItemHeight : oddItemHeight;
    }
    return h;
  }

  private _onRenderCell = (item: ISourceRoomInfo): JSX.Element => {
    return (
      <div data-is-focusable={true} className={classNames.itemCell}>
        <DetailedRoomButton roomInfo={item} onClickFn={this.onClickItem}/>
      </div>
    );
  };

  private onClickItem = (clickedItem: DetailedRoomButton) => {
    console.log('item clicked!');
    if (clickedItem === this.state.selectedItem) {
      this.setState({ selectedItem: null });
      clickedItem.setState({ selected: !clickedItem.state.selected });
    } else if (this.state.selectedItem === null) {
      this.setState({ selectedItem: clickedItem });
      clickedItem.setState({ selected: !clickedItem.state.selected });
    } else {
      this.state.selectedItem.setState({ selected: false });
      this.setState({ selectedItem: clickedItem });
      clickedItem.setState({ selected: !clickedItem.state.selected });
    }

    //clickedItem.state.selected = !clickedItem.state.selected;
  }
}
