import type { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  appId: 'org.opencells.app',
  appName: 'OpenCells',
  webDir: 'dist',
  android: {
    // WebView debugging follows the build type (on for debug, off for release)
    allowMixedContent: false,
  },
  plugins: {},
};

export default config;
