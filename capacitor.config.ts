import type { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  appId: 'com.zetl.petak',
  appName: 'Petak',
  webDir: 'dist',
  android: {
    // WebView debugging follows the build type (on for debug, off for release)
    allowMixedContent: false,
  },
  plugins: {},
};

export default config;
