/** Typed wrappers for the app's own Capacitor plugins (Java sources in android/app/src/main/java). */
import { Capacitor, registerPlugin } from '@capacitor/core';

export interface DeviceInfoResult {
  totalMem: number; // bytes
  availMem: number;
  cores: number;
  socManufacturer: string;
  socModel: string;
  hardware: string;
  sdkInt: number;
  isLowRamDevice: boolean;
  maxRefreshRate: number;
  abis: string[];
}
export interface DeviceInfoPlugin {
  getInfo(): Promise<DeviceInfoResult>;
  /** Ask the window for the highest refresh rate mode (90/120 Hz) or default. */
  setHighRefreshRate(opts: { enabled: boolean }): Promise<{ refreshRate: number }>;
}

export interface PickedDoc {
  uri: string;
  name: string;
  size: number;
  /** absolute path of a private cached copy, readable via Capacitor.convertFileSrc */
  path: string;
}
export interface DocumentsPlugin {
  open(opts: { mimeTypes: string[] }): Promise<PickedDoc>;
  /** Re-read a previously granted URI (persistable permission). */
  read(opts: { uri: string }): Promise<PickedDoc>;
  write(opts: { uri: string; data: string /* base64 */ }): Promise<void>;
  create(opts: { name: string; mimeType: string; data: string }): Promise<{ uri: string; name: string }>;
  share(opts: { name: string; mimeType: string; data: string }): Promise<void>;
}

export const isNative = () => Capacitor.isNativePlatform();
export const DeviceInfo = registerPlugin<DeviceInfoPlugin>('DeviceInfo');
export const Documents = registerPlugin<DocumentsPlugin>('Documents');
