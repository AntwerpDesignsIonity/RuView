# release_bins/

This directory holds pre-built ESP-IDF binaries for `provision.py` to flash.

**It is intentionally empty in the repo.** The Apr 2026 pre-built blob was
removed because it ignored the NVS `node_id` override and caused every
provisioned node to report as `node_id=1`, colliding with real node 1.

## Recommended (verified working) — flash via PlatformIO, NVS-only provision

```bash
cd ionity && pio run -e esp32s3_n16r8 -t upload --upload-port <PORT>
python firmware/esp32-csi-node/provision.py --port <PORT> \
    --ssid <SSID> --password <PASS> --target-ip <HUB_IP> --node-id <N> \
    --no-firmware
```

The PlatformIO firmware in `ionity/` correctly honors the NVS `node_id` key
via the Arduino `Preferences` library.

## Alternative — rebuild ESP-IDF firmware here

```bash
cd firmware/esp32-csi-node
source ~/esp/esp-idf/export.sh   # ESP-IDF v5.4+
idf.py set-target esp32s3
idf.py build
# Copy build/esp32-csi-node.bin, build/bootloader/bootloader.bin,
# build/partition_table/partition-table.bin, build/ota_data_initial.bin
# into release_bins/ and run provision.py without --no-firmware.
```

`provision.py` will warn and refuse to flash if the bin in `release_bins/` is
more than ~1 hour older than the `main/` source. Pass `--allow-stale-firmware`
to override (not recommended).

## Quarantined directory

`release_bins.stale-2026-04-06/` holds the broken bin for forensic reference
only. **Do not flash it.** It is excluded from packaging.
