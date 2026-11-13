# OpenCells APKs

| File | Use |
|---|---|
| `OpenCells-1.0.0-release.apk` | Normal install. Minified (R8), signed. Package `org.opencells.app`. |
| `OpenCells-1.0.0-debug.apk` | For testing. WebView debugging is on (`chrome://inspect`). Package `org.opencells.app.debug`, so it installs next to the release build. |

Both are universal APKs: Android 8.0+ on arm64-v8a, armeabi-v7a and x86_64 (no native code).
Check the download with `sha256sum -c SHA256SUMS.txt`.

Release signing certificate SHA-256: `c3:35:5b:e1:d0:8f:07:9c:c4:84:82:67:58:3a:76:e1:0a:64:b8:4f:db:35:5a:58:cf:55:54:3e:67:21:e7:52`

To install, open the APK on the phone and allow "Install unknown apps" for the app you opened it with.
