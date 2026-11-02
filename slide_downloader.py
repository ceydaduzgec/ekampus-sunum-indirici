import requests
import os
import re
import sys
import shutil
from PIL import Image
from io import BytesIO
import time
from tqdm import tqdm
import cairosvg
import argparse
from pptx import Presentation
from pptx.util import Inches
from collections import namedtuple
from urllib.parse import urlparse

# One line of an --input file: "Presentation name: URL"
PresentationEntry = namedtuple('PresentationEntry', ['line_number', 'name', 'url'])

INPUT_SEPARATOR = ': '


class InputFileError(Exception):
    """Raised when the --input file cannot be read or has an invalid format."""


def safe_filename(name):
    """Replace characters that are not allowed in file names."""
    return re.sub(r'[<>:"/\\|?*]', '_', name).strip()


def parse_input_file(input_path):
    """
    Parse an input file where each non-empty line is "Presentation name: URL".

    Each line describes one presentation (e.g. one week's lecture); the URL is
    the same base URL that would be passed on the command line. The line is
    split on the first ": " so the colon in "https://" is not the separator.

    Args:
        input_path (str): Path to the .txt input file

    Returns:
        list: PresentationEntry records in file order

    Raises:
        InputFileError: If the file cannot be read or a line is malformed
    """
    try:
        with open(input_path, 'r', encoding='utf-8-sig') as f:
            lines = f.read().splitlines()
    except FileNotFoundError:
        raise InputFileError(f"Input file not found: {input_path}")
    except IsADirectoryError:
        raise InputFileError(f"Input path is a directory, not a file: {input_path}")
    except PermissionError:
        raise InputFileError(f"Permission denied reading input file: {input_path}")
    except UnicodeDecodeError as e:
        raise InputFileError(f"Input file is not valid UTF-8 text: {input_path} ({e})")
    except OSError as e:
        raise InputFileError(f"Could not read input file {input_path}: {e}")

    entries = []
    errors = []
    seen_files = {}
    for line_number, raw_line in enumerate(lines, start=1):
        line = raw_line.strip()
        if not line:
            continue

        if INPUT_SEPARATOR not in line:
            errors.append(f"line {line_number}: missing '{INPUT_SEPARATOR}' separator "
                          f"(expected 'Presentation name: URL'): {line}")
            continue

        name, url = line.split(INPUT_SEPARATOR, 1)
        name = name.strip()
        url = url.strip()

        if not safe_filename(name):
            errors.append(f"line {line_number}: presentation name is empty: {line}")
            continue

        parsed = urlparse(url)
        if parsed.scheme not in ('http', 'https') or not parsed.netloc:
            errors.append(f"line {line_number}: invalid URL (must start with http:// or https://): {url}")
            continue

        # Each name becomes a .pptx file, so names must not collide
        file_key = safe_filename(name).lower()
        if file_key in seen_files:
            errors.append(f"line {line_number}: duplicate presentation name "
                          f"(same as line {seen_files[file_key]}): {name}")
            continue
        seen_files[file_key] = line_number

        entries.append(PresentationEntry(line_number, name, url))

    if errors:
        raise InputFileError(f"Invalid format in input file {input_path}:\n  " + "\n  ".join(errors))
    if not entries:
        raise InputFileError(f"Input file contains no presentations: {input_path}")

    return entries


def download_and_convert_slides(base_url, output_dir, start_slide=1, max_slides=150, format="jpeg"):
    """
    Download SVG slides and convert them to the specified format.

    Args:
        base_url (str): Base URL for slides (without slide number and extension)
        output_dir (str): Directory to save the converted images
        start_slide (int): Starting slide number
        max_slides (int): Maximum number of slides to attempt downloading
        format (str): Output format (jpeg, png)

    Returns:
        list: Paths to successfully downloaded and converted images
    """
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Prepare URL (ensure it ends with 'slide' without number or extension)
    if not base_url.endswith('slide'):
        if base_url.endswith('/'):
            base_url = base_url + 'slide'
        elif base_url.endswith('.svg'):
            # Extract base URL from a complete URL like slide1.svg
            base_url = base_url.rsplit('slide', 1)[0] + 'slide'
        else:
            base_url = base_url + '/slide'

    downloaded_files = []
    consecutive_failures = 0
    max_consecutive_failures = 3

    print(f"Starting download from: {base_url}1.svg")

    # Use tqdm for a progress bar
    progress_bar = tqdm(range(start_slide, start_slide + max_slides))

    for slide_num in progress_bar:
        if consecutive_failures >= max_consecutive_failures:
            print(f"\nStopping after {max_consecutive_failures} consecutive failures")
            break

        slide_url = f"{base_url}{slide_num}.svg"
        output_file = os.path.join(output_dir, f"slide{slide_num}.{format}")

        progress_bar.set_description(f"Downloading slide {slide_num}")

        try:
            # Download SVG
            response = requests.get(slide_url, timeout=10)
            response.raise_for_status()  # Raise exception for 404s or other errors

            # Convert SVG to desired format
            if format == "jpeg" or format == "jpg":
                # Use cairosvg to convert SVG to PNG
                png_data = cairosvg.svg2png(bytestring=response.content)
                # Use PIL to convert PNG to JPEG
                image = Image.open(BytesIO(png_data))
                rgb_image = image.convert('RGB')  # Convert to RGB for JPEG
                rgb_image.save(output_file, quality=90)
            elif format == "png":
                png_data = cairosvg.svg2png(bytestring=response.content)
                with open(output_file, 'wb') as f:
                    f.write(png_data)
            else:
                raise ValueError(f"Unsupported format: {format}")

            downloaded_files.append(output_file)
            consecutive_failures = 0

        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 404:
                consecutive_failures += 1
                progress_bar.set_description(f"Slide {slide_num} not found")
            else:
                print(f"\nError downloading {slide_url}: {e}")
                consecutive_failures += 1
        except Exception as e:
            print(f"\nError processing {slide_url}: {e}")
            consecutive_failures += 1

        # Small delay to avoid overwhelming the server
        time.sleep(0.5)

    print(f"\nSuccessfully downloaded {len(downloaded_files)} slides")
    return downloaded_files

def create_presentation(image_files, output_file):
    """
    Create a PowerPoint presentation with the downloaded images.

    Args:
        image_files (list): List of image file paths
        output_file (str): Path to save the PowerPoint presentation
    """
    prs = Presentation()

    # Set slide dimensions (16:9 aspect ratio)
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    for image_file in tqdm(image_files, desc="Creating presentation"):
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank slide

        # Calculate image dimensions to fit the slide
        img = Image.open(image_file)
        width, height = img.size

        # Calculate aspect ratio
        img_ratio = width / height
        slide_ratio = prs.slide_width / prs.slide_height

        if img_ratio > slide_ratio:
            # Image is wider than slide
            pic_width = prs.slide_width
            pic_height = pic_width / img_ratio
            left = 0
            top = (prs.slide_height - pic_height) / 2
        else:
            # Image is taller than slide
            pic_height = prs.slide_height
            pic_width = pic_height * img_ratio
            top = 0
            left = (prs.slide_width - pic_width) / 2

        # Add picture to slide
        slide.shapes.add_picture(image_file, left, top, width=pic_width, height=pic_height)

    prs.save(output_file)
    print(f"Presentation saved to {output_file}")

def cleanup_temp_files(output_dir):
    """
    Delete the temporary directory with downloaded slide images.

    Args:
        output_dir (str): Directory to remove
    """
    try:
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
            print(f"Cleaned up temporary files: {output_dir} directory deleted")
    except Exception as e:
        print(f"Warning: Could not delete temporary directory {output_dir}: {e}")

def build_presentation(base_url, output_dir, presentation_file, args):
    """
    Download one presentation's slides and save them as a PowerPoint file.

    Returns:
        bool: True if the presentation was created
    """
    # Download and convert slides
    image_files = download_and_convert_slides(
        base_url,
        output_dir,
        args.start,
        args.max,
        args.format
    )

    if not image_files:
        return False

    # Create presentation
    create_presentation(image_files, presentation_file)

    # Clean up temporary files unless --keep-images is specified
    if not args.keep_images:
        cleanup_temp_files(output_dir)
    return True

def main():
    parser = argparse.ArgumentParser(description='Download slides and create a presentation')
    parser.add_argument('base_url', nargs='?', help='Base URL for slides (e.g., https://example.com/presentation/id/svgs/slide)')
    parser.add_argument('--input', metavar='FILE', help='Text file with one "Presentation name: URL" per line; '
                        'creates "<name>.pptx" for each line instead of using base_url')
    parser.add_argument('--output-dir', default='slides', help='Directory to save the images (default: slides)')
    parser.add_argument('--start', type=int, default=1, help='Starting slide number (default: 1)')
    parser.add_argument('--max', type=int, default=150, help='Maximum number of slides to download (default: 150)')
    parser.add_argument('--format', choices=['jpeg', 'png'], default='jpeg', help='Output image format (default: jpeg)')
    parser.add_argument('--presentation', default='presentation.pptx', help='Output PowerPoint file (default: presentation.pptx); '
                        'ignored with --input')
    parser.add_argument('--keep-images', action='store_true', help='Keep the image files after creating the presentation')

    args = parser.parse_args()

    if args.base_url and args.input:
        parser.error('use either base_url or --input, not both')
    if not args.base_url and not args.input:
        parser.error('either base_url or --input is required')

    # Read the input file before doing any work so format errors are reported up front
    entries = None
    if args.input:
        try:
            entries = parse_input_file(args.input)
        except InputFileError as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(1)

    print("Slide Downloader and Presentation Creator")
    print("=========================================")

    # Generate unique output_dir if default is used
    if args.output_dir == 'slides':
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        args.output_dir = f"slides_{timestamp}"

    if entries is None:
        build_presentation(args.base_url, args.output_dir, args.presentation, args)
        return

    # --input: run the single-presentation flow once per line
    if args.presentation != 'presentation.pptx':
        print("Warning: --presentation is ignored with --input; each presentation is saved as '<name>.pptx'")

    print(f"Read {len(entries)} presentations from {args.input}")
    failed = []
    for index, entry in enumerate(entries, start=1):
        print(f"\n[{index}/{len(entries)}] {entry.name}")
        file_name = safe_filename(entry.name)
        output_dir = os.path.join(args.output_dir, file_name)
        presentation_file = f"{file_name}.pptx"
        try:
            created = build_presentation(entry.url, output_dir, presentation_file, args)
        except Exception as e:
            print(f"Error creating '{entry.name}': {e}")
            created = False
        if not created:
            failed.append(entry)

    # Remove the parent image directory, including folders left by failed presentations
    if not args.keep_images:
        cleanup_temp_files(args.output_dir)

    print(f"\nCreated {len(entries) - len(failed)} of {len(entries)} presentations")
    if failed:
        print("Failed presentations:")
        for entry in failed:
            print(f"  line {entry.line_number}: {entry.name} ({entry.url})")
        sys.exit(1)

if __name__ == "__main__":
    main()