package com.zetl.petak;

import android.os.Bundle;
import android.webkit.WebResourceRequest;
import android.webkit.WebResourceResponse;
import android.webkit.WebView;
import com.getcapacitor.BridgeActivity;
import com.getcapacitor.BridgeWebViewClient;
import java.util.HashMap;
import java.util.Map;

/**
 * Entry point. Registers the app's own native plugins:
 *  - DeviceInfo: RAM / CPU cores / SoC / refresh rate / thermal status for performance tiering
 *  - Documents:  Storage Access Framework open / save-in-place / create / share
 *
 * The app's own files are served with COOP/COEP headers so the page can be
 * cross-origin isolated, which enables SharedArrayBuffer between the UI and the
 * workers. If the WebView does not honour them, the app falls back to messages.
 */
public class MainActivity extends BridgeActivity {
    @Override
    public void onCreate(Bundle savedInstanceState) {
        registerPlugin(DeviceInfoPlugin.class);
        registerPlugin(DocumentsPlugin.class);
        super.onCreate(savedInstanceState);
        try {
            bridge.setWebViewClient(new BridgeWebViewClient(bridge) {
                @Override
                public WebResourceResponse shouldInterceptRequest(WebView view, WebResourceRequest request) {
                    WebResourceResponse r = super.shouldInterceptRequest(view, request);
                    if (r != null && bridge.getHost().equals(request.getUrl().getHost())) {
                        Map<String, String> h = r.getResponseHeaders() == null ? new HashMap<>() : new HashMap<>(r.getResponseHeaders());
                        h.put("Cross-Origin-Opener-Policy", "same-origin");
                        h.put("Cross-Origin-Embedder-Policy", "credentialless");
                        r.setResponseHeaders(h);
                    }
                    return r;
                }
            });
        } catch (RuntimeException e) {
            // keep the default client: SharedArrayBuffer is optional
        }
    }
}
