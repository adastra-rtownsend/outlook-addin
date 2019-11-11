const _defaultSettings = {
  showWelcomeScreen: 2, // 1 = never, 2 = next time, 3 = always
  useSampleData: false,
  apiBasePath: 'https://ache-bridge-api.herokuapp.com',
}

export const WELCOME_SCREEN_SETTTING = 'adastra.demo.showWelcomeScreen';
export const DEMO_DATA_SETTING = 'adastra.demo.useSampleData';
export const API_PATH_SETTING  = 'adastra.demo.apiBasePath';
export const SELECTED_ROOM_SETTING = 'adastra.demo.selectedRoom';

export function getDefaultSettings() {
  return _defaultSettings;
};
