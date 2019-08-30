// derivative of reference code  for creating example data iwth  office-ui-fabric-react.

const DATA = {    
  roomNames: [
                '1st Floor - Engineering Conference Room', 
                '2nd Floor - Conference Room (Services)', 
                '2nd Floor - Conference Room (Sales)'
              ],
  available: ['true', 'false']  
};

export interface IExampleItem {
  key: string;
  roomName: string;
  available: boolean;
  capacity: number;
};

export function createListItems(count: number, startIndex: number = 0): IExampleItem[] {
  return Array.apply(null, Array(count)).map((item: number, index: number) => {

    return {
      key: 'item-' + (index + startIndex) + (item === undefined ? '-empty' : '-not empty'),
      roomName: _randWord(DATA.roomNames),
      available: _randWord(DATA.available),
    };
  });
};

function _randWord(array: string[]): string {
  const index = Math.floor(Math.random() * array.length);
  return array[index];
};
