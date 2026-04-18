## AEDI-S MAUI client

Cross-platform shell that connects to an AEDI-S hub over WebSocket and shows
live sensing readings (presence, motion, persons, HR, BR, PSO localization).

### Targets

| Platform | Target framework | Workload required |
|----------|------------------|-------------------|
| Android  | `net10.0-android` | `maui-android` |
| iOS      | `net10.0-ios`     | `maui-ios` (macOS host) |
| macOS    | `net10.0-maccatalyst` | `maui-maccatalyst` (macOS host) |
| Windows  | `net10.0-windows10.0.19041.0` | `maui-windows` (Windows host) |

### Build

```bash
# auto-pick installed targets
./build.sh

# or explicit
./build.sh release android
dotnet build -c Release -f net10.0-android
```

### Install to device (Android)

```bash
dotnet build -c Release -f net10.0-android -t:Install
```

APK lives at `bin/Release/net10.0-android/today.ionity.aedi-Signed.apk`.

### Connect

On first launch, enter the hub URL e.g.
`ws://192.168.124.7:3001/ws/sensing` and tap **Connect**.
The app also queries `http://<host>:3000/api/v1/localization/person`
every ~10 ticks to show PSO position.
