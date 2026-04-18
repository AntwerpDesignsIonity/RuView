# Hardware Reference — ESP32 Boards in AEDI-S

> Verified pinouts and quirks for the ESP32 boards used as CSI sensing nodes.
> When the upstream Waveshare schematic is wrong or ambiguous, the
> CircuitPython board file is the authoritative source.

## Supported Boards

| node-id | Board                                  | Chip       | Flash | PSRAM | LCD                  | Touch    | PlatformIO env        |
|--------:|----------------------------------------|------------|------:|------:|----------------------|----------|-----------------------|
|       1 | Generic ESP32-S3-DevKitC-1 (N16R8)     | ESP32-S3   |  16MB |   8MB | —                    | —        | `esp32s3_n16r8`       |
|       2 | Generic ESP32-S3-DevKitC-1 (N16R8)     | ESP32-S3   |  16MB |   8MB | —                    | —        | `esp32s3_n16r8`       |
|       3 | Generic ESP32-S3-DevKitC-1 (N16R8)     | ESP32-S3   |  16MB |   8MB | —                    | —        | `esp32s3_n16r8`       |
|       4 | Waveshare ESP32-S3-Touch-LCD-2         | ESP32-S3   |  16MB |   8MB | ST7789 240×320 IPS   | CST816D  | `esp32s3_touch_lcd_2` |
|       5 | Generic ESP32-S3-DevKitC-1 (N16R8)     | ESP32-S3   |  16MB |   8MB | —                    | —        | `esp32s3_n16r8`       |
|       6 | Waveshare ESP32-S3-LCD-1.47            | ESP32-S3   |  16MB |   8MB | ST7789V3 172×320     | —        | `esp32s3_lcd_1_47`    |

The current per-MAC mapping lives in
[firmware/esp32-csi-node/nodes.yaml](../../firmware/esp32-csi-node/nodes.yaml).

## Waveshare ESP32-S3-Touch-LCD-2 (240×320)

The pin map in the official Waveshare wiki and PDF schematic is **incorrect**
for several signals. The CircuitPython board file
([`ports/espressif/boards/waveshare_esp32_s3_touch_lcd_2/pins.c`][cp-touch-lcd-2])
is correct and matches the board we have.

| Signal      | GPIO | Notes                                            |
|-------------|-----:|--------------------------------------------------|
| `LCD_SCLK`  |   39 | SPI clock                                        |
| `LCD_MOSI`  |   38 | SPI data out                                     |
| `LCD_MISO`  |   40 | SPI data in (rarely used by ST7789)              |
| `LCD_CS`    |   45 | Chip select  ← Waveshare schematic says 42       |
| `LCD_DC`    |   42 | Data/command  ← Waveshare schematic says 41      |
| `LCD_RST`   |    0 | Reset  ← Waveshare schematic says 40             |
| `LCD_BL`    |    1 | Backlight  ← Waveshare schematic says 15 (ADC)   |
| `TOUCH_SDA` |   48 | I²C data (CST816D)                               |
| `TOUCH_SCL` |   47 | I²C clock                                        |
| `TOUCH_INT` |   46 | Touch interrupt                                  |

Source of truth: [`ionity/src/lcd_st7789_240x320.h`](../../ionity/src/lcd_st7789_240x320.h)

[cp-touch-lcd-2]: https://github.com/adafruit/circuitpython/blob/main/ports/espressif/boards/waveshare_esp32_s3_touch_lcd_2/pins.c

## Waveshare ESP32-S3-LCD-1.47 (172×320)

Pin map verified from the same CircuitPython source — these match the
schematic for this board.

See [`ionity/src/lcd_st7789_147.h`](../../ionity/src/lcd_st7789_147.h) for
the canonical values used by the firmware.

## ESP32-S3 USB Modes

Two distinct USB modes appear on different `lsusb` lines:

| `lsusb` ID  | Mode                  | Reset behaviour                      | Notes                                  |
|-------------|-----------------------|--------------------------------------|----------------------------------------|
| `303a:1001` | USB-Serial-JTAG (ROM) | esptool can flash directly           | Always present on bare ROM; preferred  |
| `303a:4001` | Arduino USB-CDC       | DTR/RTS reset is **broken** on Linux | esptool fails with `BrokenPipeError`   |

If you see `303a:4001`, hold BOOT and replug — the board enters ROM and
appears as `303a:1001`, after which `pio run -t upload` works without
touching the BOOT button.

## Unsupported

ESP32 (original) and ESP32-C3 are single-core and **cannot** run the CSI
DSP pipeline. The provisioner refuses these chips.
