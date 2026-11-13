# Petak APKs

Release builds are published on the repository's **Releases** page by GitHub Actions
(workflow `Petak Android APK`, triggered by pushing a `v*` tag). Files are named
`Petak-v<version>.apk` (release) and `Petak-v<version>-debug.apk` (debug).

| File | Use |
|---|---|
| `Petak-v1.1.0.apk` | Normal install. Minified (R8), signed. Package `com.zetl.petak`. |
| `Petak-v1.1.0-arm64-v8a.apk`, `-armeabi-v7a.apk`, `-x86_64.apk` | Same app, one per ABI (identical content: Petak has no native code). |
| `Petak-v1.1.0-debug.apk` | For testing. WebView debugging is on (`chrome://inspect`). Package `com.zetl.petak.debug`, installs next to the release build. |
| `SIZE.md` | Size of each build and module. |

All APKs run on Android 8.0+. Without the repository's signing secrets, CI signs release APKs with its debug key.
Verify a download with `sha256sum -c SHA256SUMS.txt`.

Petak replaces the earlier "OpenCells" builds (package `org.opencells.app`). Android treats it as a new app,
so both can be installed side by side; uninstall the old one when you no longer need it.
