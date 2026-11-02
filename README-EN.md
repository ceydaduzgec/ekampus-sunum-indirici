# Slide Downloader and Presentation Creator

This Python script downloads SVG slides from a specified URL, converts them to JPEG or PNG format, and creates a PowerPoint presentation.

## Setup Instructions

1. **Install Required Dependencies**

   First, ensure you have Python 3.6+ installed on your system. Then install the required packages:

   ```bash
   pip install -r requirements.txt
   ```

   Note: For CairoSVG to work properly, you may need to install additional system dependencies:
   
   - **On Ubuntu/Debian**:
     ```bash
     sudo apt-get install libcairo2-dev libpango1.0-dev libgdk-pixbuf2.0-dev libffi-dev shared-mime-info
     ```
   
   - **On macOS** (using Homebrew):
     ```bash
     brew install cairo pango gdk-pixbuf libffi
     ```
   
   - **On Windows**:
     It's recommended to use the pre-built binary wheels that pip will install automatically.

2. **Get the URL of the slide**
   
   Right-click on a slide in your browser and select "Copy Link Address" or "Copy Link" to get the URL of the slide.

<div align="center">
<img src="images/copy_url.png" alt="How to copy slide URL" width="80%" style="border: 1px solid #ddd; border-radius: 5px;" />
</div>

3. **Run the Script**

   Use the script to download slides and create a presentation:

   ```bash
   python slide_downloader.py https://vkscalelite.anadolu.edu.tr/presentation/633fc11f5c4832a9db123d72eb67831a012d0710-1741272327834/presentation/091f41e4a0d9eaa5c28fad9596121aa1ed2b315c-1741272345267/svgs/slide1.svg
   ```

   The script will:
   - Download all slides (starting from slide1.svg)
   - Convert them to JPEG format
   - Create a PowerPoint presentation with all slides
   - Automatically delete the temporary image files

## Command-Line Options

```
usage: slide_downloader.py [-h] [--input FILE] [--output-dir OUTPUT_DIR] [--start START]
                          [--max MAX] [--format {jpeg,png}] [--presentation PRESENTATION]
                          [--keep-images]
                          [base_url]

Download slides and create a presentation

positional arguments:
  base_url              Base URL for slides (e.g., https://example.com/presentation/id/svgs/slide)
                        (required unless --input is used)

optional arguments:
  -h, --help                   Show this help message and exit
  --input FILE                 Text file with one "Presentation name: URL" per line; creates a
                               separate "<name>.pptx" for each line instead of using base_url
  --output-dir OUTPUT_DIR      Directory to save the images (default: slides)
  --start START                Starting slide number (default: 1)
  --max MAX                    Maximum number of slides to download (default: 100)
  --format {jpeg,png}          Output image format (default: jpeg)
  --presentation PRESENTATION  Output PowerPoint file (default: presentation.pptx); ignored with --input
  --keep-images                Keep the temporary image files after creating the presentation
```

## Examples

1. **Basic usage**:
   ```bash
   python slide_downloader.py https://vkscalelite.anadolu.edu.tr/presentation/633fc11f5c4832a9db123d72eb67831a012d0710-1741272327834/presentation/091f41e4a0d9eaa5c28fad9596121aa1ed2b315c-1741272345267/svgs/slide
   ```

2. **Specify range of slides**:
   ```bash
   python slide_downloader.py https://vkscalelite.anadolu.edu.tr/presentation/633fc11f5c4832a9db123d72eb67831a012d0710-1741272327834/presentation/091f41e4a0d9eaa5c28fad9596121aa1ed2b315c-1741272345267/svgs/slide --start 5 --max 20
   ```

3. **Use PNG format**:
   ```bash
   python slide_downloader.py https://vkscalelite.anadolu.edu.tr/presentation/633fc11f5c4832a9db123d72eb67831a012d0710-1741272327834/presentation/091f41e4a0d9eaa5c28fad9596121aa1ed2b315c-1741272345267/svgs/slide --format png
   ```

4. **Custom output location**:
   ```bash
   python slide_downloader.py https://vkscalelite.anadolu.edu.tr/presentation/633fc11f5c4832a9db123d72eb67831a012d0710-1741272327834/presentation/091f41e4a0d9eaa5c28fad9596121aa1ed2b315c-1741272345267/svgs/slide --output-dir my_slides --presentation my_lecture.pptx
   ```

5. **Keep temporary image files**:
   ```bash
   python slide_downloader.py https://vkscalelite.anadolu.edu.tr/presentation/633fc11f5c4832a9db123d72eb67831a012d0710-1741272327834/presentation/091f41e4a0d9eaa5c28fad9596121aa1ed2b315c-1741272345267/svgs/slide --keep-images
   ```

6. **Download several weeks at once (`--input`)**:
   ```bash
   python3 slide_downloader.py --input weeks.txt
   ```
   See [Batch Download from an Input File](#batch-download-from-an-input-file) below for details.

## Batch Download from an Input File

Instead of running the command once per week, you can list every week in a `.txt` file and download them all at once with `--input`. Each line in the file is **one week's presentation**, in the form `Presentation name: URL`. The URL is the start link of that week's presentation, the same link you would pass as `base_url` for a single presentation.

```text
2024-2025 Spring_Week 1: https://vkscalelite.anadolu.edu.tr/presentation/.../svgs/slide1.svg
2024-2025 Spring_Week 2: https://vkscalelite.anadolu.edu.tr/presentation/.../svgs/slide1.svg
```

```bash
python3 slide_downloader.py --input weeks.txt
```

This behaves as if you ran the existing command once for each line. For example, the first line is equivalent to:

```bash
python3 slide_downloader.py https://vkscalelite.anadolu.edu.tr/presentation/633fc11f5c4832a9db123d72eb67831a012d0710-1741272327834/presentation/091f41e4a0d9eaa5c28fad9596121aa1ed2b315c-1741272345267/svgs/slide1.svg --presentation "2024-2025 Spring_Week 1.pptx"
```

- Each line produces a presentation named after it (e.g. `2024-2025 Spring_Week 1.pptx`), saved in the directory you run the command from. Characters that aren't allowed in file names (`/ \ : * ? " < > |`) are replaced with `_`.
- Lines are processed in the order they appear in the file.
- The name and URL are split on the **first** `: ` (colon + space). The colon in `https://` is not treated as the separator, so don't use `: ` inside a presentation name.
- Blank lines are skipped. The file must be UTF-8. The same name can't be used twice.
- `--start`, `--max`, `--format` and `--keep-images` apply to each presentation separately. With `--keep-images`, images are kept under `--output-dir` in a separate folder per presentation. `--presentation` is ignored in this mode.
- `base_url` and `--input` cannot be used together.
- If the file can't be read or a line is malformed, an error listing the offending line numbers is shown before any download starts.
- If one week fails to download, the remaining weeks continue. Failed lines are listed at the end.

## Tests

The tests serve slides from a local HTTP server, so no internet access is needed. The dependencies from the setup steps (including CairoSVG's system dependencies) must be installed. Run this from the project directory:

```bash
python3 -m unittest discover tests
```

## Troubleshooting

1. **SVG Conversion Issues**:
   If CairoSVG fails to convert some slides, it might be due to malformed SVG content. Try downloading the SVGs separately and use a different converter.

2. **404 Errors**:
   The script will automatically stop after 3 consecutive 404 errors, assuming it has reached the end of available slides.

3. **Rate Limiting**:
   If you're getting connection errors, try increasing the delay between downloads by modifying the `time.sleep()` value in the script.