const _defaultSettings = {
  showIntro: 1, // 0 = never, 1 = next time, 2 = always
  useSampleData: true,
  apiBasePath: 'http://localhost:2999',
}

export function getDefaultSettings() {
  return _defaultSettings;
};
