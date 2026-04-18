# Troubleshooting — AEDI-S Stack

Quick reference for the most common failure modes. If your problem isn't
here, check `logs/sensing-server.log` and `logs/aedi.log` first.

## Stack won't start

### `cargo build` fails with linker errors
The Pi runs out of memory linking the release binary. Use a swap file or
build on a workstation and copy the binary across:

```bash
# On the Pi
sudo fallocate -l 4G /swapfile && sudo mkswap /swapfile && sudo swapon /swapfile
```

### `Server did not respond at http://localhost:3000/health within 15s`
The Rust binary started but the UI dir is wrong, or the UDP/TCP ports are
already taken. Check:
```bash
ss -tlnp | grep -E '3000|3001'   # who owns the ports
tail -50 logs/sensing-server.log # actual error
ls ui/index.html                 # path passed via --ui-path
```

### Disk full mid-build
The Pi root partition fills up fast. PIO will crash with
`OSError: [Errno 28] No space left on device` from `SCons/dblite.py`. Fix:
```bash
make clean-pio   # frees ~1-2 GB safely
df -h /
```
If still tight: `cargo clean` in `rust-port/wifi-densepose-rs/`, then delete
old `.pio/build/<env>` directories you no longer need.

## ESP32 won't flash

### `BrokenPipeError` on TIOCMSET
The board enumerated as `303a:4001` (Arduino USB-CDC) — DTR/RTS reset is
broken on Linux for this mode.

**Fix:** Hold the BOOT button, replug USB, release BOOT. Confirm with
`lsusb | grep 303a` — should now show `303a:1001` (USB-Serial-JTAG ROM).
esptool can flash this directly.

### Chip stuck in download mode after flash
esptool defaulted to `--after=no_reset`. Reset manually with the RST
button, or re-run with `--after=hard_reset`.

### `pio` not found
On the Pi there's no system-wide `pio`:
```bash
~/.platformio/penv/bin/pio run -e <env> -t upload
```
Or alias it: `alias pio=~/.platformio/penv/bin/pio`.

## Wrong / blank LCD

The Waveshare schematic for ESP32-S3-Touch-LCD-2 has wrong pins for
`DC`/`CS`/`RST`/`BL`. Use [docs/hardware/esp32-boards.md](hardware/esp32-boards.md)
for the verified map. The firmware in
[`ionity/src/lcd_st7789_240x320.h`](../ionity/src/lcd_st7789_240x320.h) is
already correct — re-flash with the right env, don't tweak the pins.

## Node is provisioned but no data on the dashboard

1. Confirm WiFi: read serial at 460800 baud, look for `[LED] WiFi connected!`.
2. Confirm UDP target IP: `[CSI] streaming to <hub-ip>:5005`.
3. Hub firewall: `sudo ufw status` — port 5005/udp must be allowed.
4. CSI bridge running: `tail -20 logs/csi-bridge.log`.
5. Re-provision quickly without re-flash:
   ```bash
   python firmware/esp32-csi-node/identify-node.py --port /dev/ttyACM0 --provision
   ```

## Tests / proof failing

### `python v1/data/proof/verify.py` says hash mismatch
Pre-existing drift from numpy/scipy version variance.
See [docs/adr/ADR-044-proof-hash-stability.md](adr/ADR-044-proof-hash-stability.md)
(if you regenerated the hash) or pin numpy/scipy in `requirements.txt`.

### `ruvsense::field_model::test_estimate_occupancy_noise_only` fails
Pre-existing failure on `main` (`NotCalibrated`). Verified via `git stash`.
Tracking: see GitHub issues.

### Mock mode missed bug X
Always test firmware changes against **real WiFi CSI**, not the mock
generator. The mock missed the Kconfig threshold bug in v0.7.

## Provisioning workflow

For new boards, the smoothest path is:

```bash
# 1. Identify (autodetect what's plugged in)
python firmware/esp32-csi-node/identify-node.py

# 2. If MAC is known: re-provision (no flash)
python firmware/esp32-csi-node/identify-node.py --port /dev/ttyACM0 --provision

# 3. If MAC is unknown: claim a free node-id slot
python firmware/esp32-csi-node/identify-node.py --port /dev/ttyACM0 \
    --node-id 3 --flash      # writes MAC into nodes.yaml + flashes + provisions

# 4. Verify
python firmware/esp32-csi-node/identify-node.py
# expect: ==  /dev/ttyACM0  mac=...  node-id=3  env=esp32s3_n16r8
```

## More

- Hardware pin maps: [docs/hardware/esp32-boards.md](hardware/esp32-boards.md)
- Full user guide:    [docs/user-guide.md](user-guide.md)
- LED status codes:   [docs/led-indication.md](led-indication.md)
