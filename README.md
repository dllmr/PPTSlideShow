# PPTSlideShow

A single-file Python script that builds a looping PowerPoint slideshow from a folder of images. One slide per image, sized to fill the slide without clipping, centred on a black background, with automatic slide advance.

The output plays in both PowerPoint and LibreOffice Impress on Windows, macOS, and Linux.

## Requirements

- [uv](https://docs.astral.sh/uv/) (handles Python and dependencies automatically via PEP 723 inline metadata).

## Usage

From the folder containing your images:

```sh
uv run slideshow.py
```

This produces `slideshow.pptx` alongside your images. By default:

- 10 seconds per slide
- Loops continuously
- Images from the current folder only (subfolders ignored)
- Images are **embedded** in the PPTX (self-contained, portable) and downscaled to fit within **1920×1080** to keep file size reasonable
- No transition effect between slides

### Interactive mode

```sh
uv run slideshow.py -I
```

Prompts for:

1. Slide duration (seconds)
2. Loop on/off
3. Embed vs link images (linking keeps the PPTX small but requires the PPTX to stay at the root of the image tree when opened)
4. Fade transition on/off (fixed 0.5s duration)
5. Include images from subfolders
6. Downscale embedded images (`1080p`, `720p`, `480p`, or `none`). Only applies when embedding; linked images are never modified. Images already within the cap are embedded unchanged; larger ones are resized preserving aspect ratio and re-encoded as JPEG (for JPEG sources) or PNG (otherwise).

The chosen values are written to `slideshow.toml` in the same folder and become the defaults for subsequent runs (both interactive and non-interactive). Delete the file to reset to factory defaults.

You can also edit `slideshow.toml` directly. All keys are optional; missing keys fall back to the built-in defaults:

```toml
duration = 10.0          # seconds per slide (positive number)
loop = true              # loop when the presentation ends
embed = true             # true = embed images, false = link them
fade = false             # fade transition between slides (fixed 0.5s)
recursive = false        # include images from subfolders
scale = "1080p"          # "1080p", "720p", "480p", or "none" (embed only)
```

### Supported image formats

`.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`, `.tif`, `.tiff`, `.webp`

Files are added in alphabetical order of their path relative to the folder.

### File explorer preview

The first image is rendered onto a 16:9 black canvas and embedded as the PPTX thumbnail, so file managers that read OOXML thumbnails (Windows Explorer, GNOME Files, etc.) show a relevant preview.

## Implementation notes

A couple of quirks in LibreOffice's PPTX importer that the script works around:

- `<p:transition advTm="…"/>` is ignored unless the element has a child (e.g. `<p:cut/>`). The script always emits one.
- Loop-until-Esc (`<p:showPr loop="1">`) must live in `ppt/presProps.xml`, not `ppt/presentation.xml`. The script sets it in both for portability.
