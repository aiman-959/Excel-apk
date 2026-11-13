package com.zetl.petak;

import android.os.Bundle;
import com.getcapacitor.BridgeActivity;

/**
 * Entry point. Registers the app's own native plugins:
 *  - DeviceInfo: RAM / CPU cores / SoC / refresh rate for performance tiering
 *  - Documents:  Storage Access Framework open / save-in-place / create / share
 */
public class MainActivity extends BridgeActivity {
    @Override
    public void onCreate(Bundle savedInstanceState) {
        registerPlugin(DeviceInfoPlugin.class);
        registerPlugin(DocumentsPlugin.class);
        super.onCreate(savedInstanceState);
    }
}
