const _defaultSettings = {
  showWelcomeScreen: 2, // 1 = never, 2 = next time, 3 = always
  useSampleData: true,
  apiBasePath: 'http://localhost:3000',
}

export function getDefaultSettings() {
  return _defaultSettings;
};
