# Pixel-Perfect

Map every pixel of an image to a coloured Excel cell — and back.

## Usage

```
py main.py to-excel <image> <xlsx> [--scale N]
py main.py to-image <xlsx> <output> [--scale N]
```

### Examples

```bash
# Convert at native resolution
py main.py to-excel photo.png pixels.xlsx

# 2× upscale (e.g. 100×100 image → 200×200 grid)
py main.py to-excel photo.png pixels.xlsx --scale 2

# Half resolution (downscale)
py main.py to-excel photo.png pixels.xlsx --scale 0.5

# Edit cells in Excel, then render at 4×
py main.py to-image pixels.xlsx output.png --scale 4
```

## Notes

- **Large grids** (>500k cells) will be slow in Excel.
- **Excel limits**: 1,048,576 rows × 16,384 columns enforced.
- Scale works on both commands: `to-excel` accepts any positive float, `to-image` accepts positive integers.
