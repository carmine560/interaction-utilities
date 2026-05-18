"""Screen text recognition utilities."""

import time

import pytesseract
from PIL import Image
from PIL import ImageGrab
from PIL import ImageOps

from core_utilities.errors import TextRecognitionError


def _get_tesseract_config(text_type):
    """Return the Tesseract configuration for a supported text type."""
    config_by_type = {
        # Keep spaces inside quoted whitelist values, but not at the end. On
        # Windows, pytesseract uses shlex.split(..., posix=False), and trailing
        # spaces before closing quotes are not passed reliably.
        "integers": "-c tessedit_char_whitelist='0123456789 ,' --psm 7",
        "decimal_numbers": (
            "-c tessedit_char_whitelist='0123456789 ,.' --psm 7"
        ),
        "securities_code_column": (
            "-c tessedit_char_whitelist='0123456789"
            "ACDFGHJKLMNPRSTUWXY' --psm 6"
        ),
    }
    return config_by_type[text_type]


def _prepare_image(
    x,
    y,
    width,
    height,
    image_magnification,
    binarization_threshold,
    is_dark_theme,
):
    """Grab and preprocess an image for OCR."""
    image = ImageGrab.grab(bbox=(x, y, x + width, y + height))
    image = image.resize(
        (image_magnification * width, image_magnification * height),
        Image.Resampling.LANCZOS,
    )
    image = image.point(lambda p: 255 if p > binarization_threshold else 0)
    if is_dark_theme:
        image = ImageOps.invert(image)
    return image


def _parse_recognized_text(recognized_text, text_type):
    """Parse recognized OCR output into the expected result shape."""
    if text_type in ("integers", "decimal_numbers"):
        return [
            float(item.replace(",", "")) for item in recognized_text.split()
        ]
    if text_type == "securities_code_column":
        return recognized_text.splitlines()
    return []


def _raise_text_recognition_error(
    attempts,
    last_recognized_text,
    region,
    text_type,
):
    """Raise a structured OCR failure with retry details."""
    x, y, width, height = region
    raise TextRecognitionError(
        "OCR did not produce parseable text after "
        f"{attempts} attempts for {text_type} in region "
        f"({x}, {y}, {width}, {height}). Last output: "
        f"{last_recognized_text!r}",
        attempts=attempts,
        last_output=last_recognized_text,
        region=region,
        text_type=text_type,
    )


def recognize_text(
    x,
    y,
    width,
    height,
    index,
    image_magnification,
    binarization_threshold,
    is_dark_theme,
    text_type="integers",
    should_continue_reference=None,
    max_attempts=50,
    retry_interval_seconds=0.1,
):
    """Recognize and return text from a specified screen area."""
    config = _get_tesseract_config(text_type)

    split_text = []
    attempts = 0
    last_recognized_text = ""
    region = (x, y, width, height)
    while not split_text:
        if (
            should_continue_reference is not None
            and not should_continue_reference()
        ):
            return None
        attempts += 1
        image = _prepare_image(
            x,
            y,
            width,
            height,
            image_magnification,
            binarization_threshold,
            is_dark_theme,
        )
        recognized_text = pytesseract.image_to_string(image, config=config)
        last_recognized_text = recognized_text.strip()
        try:
            split_text = _parse_recognized_text(recognized_text, text_type)
        except ValueError:
            split_text = []

        if split_text:
            break
        if max_attempts is not None and attempts >= max_attempts:
            _raise_text_recognition_error(
                attempts,
                last_recognized_text,
                region,
                text_type,
            )
        if retry_interval_seconds > 0:
            time.sleep(retry_interval_seconds)

    if index is None:
        return split_text
    return split_text[int(index)]
