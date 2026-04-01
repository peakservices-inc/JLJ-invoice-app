#!/usr/bin/env python3
"""
Annotate scanned invoice PDFs with an OCR-derived due date and a fixed payment note.

This script:
1. Uses OCR to find the invoice date on each page.
2. Calculates the due date as invoice date + 30 days.
3. Adds a "Due Date" field to the right of the detected invoice date area.
4. Adds a fixed payment note at the bottom of each page.

Dependencies:
    pip install pymupdf pillow pytesseract

System requirement:
    Tesseract OCR must be installed and either available on PATH or passed with
    --tesseract-cmd.

Examples:
    python annotate_invoice_due_dates.py --input INPUT\\invoice.pdf --output OUTPUT\\invoice_due.pdf
    python annotate_invoice_due_dates.py
"""

from __future__ import annotations

import argparse
import io
import re
import shutil
import sys
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Callable, Iterable, List, Optional, Sequence, Tuple

import fitz
import pytesseract
from PIL import Image, ImageDraw, ImageFont, ImageOps
from pytesseract import Output


NOTE_TEXT = (
    "Note - Payment is due within 30 days from the invoice date. "
    "Accounts not settled within this period may be subject to a 5% monthly "
    "interest charge on the outstanding balance."
)

DEFAULT_INPUT_DIR = Path("INPUT")
DEFAULT_OUTPUT_DIR = Path("OUTPUT")

MONTH_PATTERN = (
    r"(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
    r"jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|"
    r"nov(?:ember)?|dec(?:ember)?)"
)

DATE_PATTERNS = [
    re.compile(r"\b\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}\b", re.IGNORECASE),
    re.compile(r"\b\d{4}[\/\-.]\d{1,2}[\/\-.]\d{1,2}\b", re.IGNORECASE),
    re.compile(r"\b%s\s+\d{1,2},?\s+\d{2,4}\b" % MONTH_PATTERN, re.IGNORECASE),
    re.compile(r"\b\d{1,2}\s+%s,?\s+\d{2,4}\b" % MONTH_PATTERN, re.IGNORECASE),
]

NUMERIC_DATE_FIXUPS = str.maketrans(
    {
        "O": "0",
        "o": "0",
        "I": "1",
        "l": "1",
        "|": "1",
        "S": "5",
    }
)

NUMERIC_DATE_FORMATS = [
    "%m/%d/%Y",
    "%m/%d/%y",
    "%m-%d-%Y",
    "%m-%d-%y",
    "%m.%d.%Y",
    "%m.%d.%y",
    "%d/%m/%Y",
    "%d/%m/%y",
    "%d-%m-%Y",
    "%d-%m-%y",
    "%d.%m.%Y",
    "%d.%m.%y",
    "%Y/%m/%d",
    "%Y-%m-%d",
    "%Y.%m.%d",
]

ALPHA_DATE_FORMATS = [
    "%b %d %Y",
    "%b %d %y",
    "%B %d %Y",
    "%B %d %y",
    "%d %b %Y",
    "%d %b %y",
    "%d %B %Y",
    "%d %B %y",
]

FONT_PATHS = {
    "regular": [
        Path(r"C:\Windows\Fonts\arial.ttf"),
        Path(r"C:\Windows\Fonts\calibri.ttf"),
        Path(r"C:\Windows\Fonts\segoeui.ttf"),
    ],
    "bold": [
        Path(r"C:\Windows\Fonts\arialbd.ttf"),
        Path(r"C:\Windows\Fonts\calibrib.ttf"),
        Path(r"C:\Windows\Fonts\segoeuib.ttf"),
    ],
}

UNSAFE_FONT_KEYWORDS = (
    "symbol",
    "wingding",
    "webding",
    "marlett",
    "emoji",
    "fluenticons",
    "mdl2",
    "assets",
    "icons",
)

OCR_CONFIGS = [
    "--oem 3 --psm 6",
    "--oem 3 --psm 11",
]


@dataclass(frozen=True)
class Box:
    x0: int
    y0: int
    x1: int
    y1: int

    @property
    def width(self) -> int:
        return max(0, self.x1 - self.x0)

    @property
    def height(self) -> int:
        return max(0, self.y1 - self.y0)

    @property
    def center_x(self) -> float:
        return (self.x0 + self.x1) / 2.0

    @property
    def center_y(self) -> float:
        return (self.y0 + self.y1) / 2.0

    def expand(self, dx: int, dy: int) -> "Box":
        return Box(self.x0 - dx, self.y0 - dy, self.x1 + dx, self.y1 + dy)


@dataclass
class OCRWord:
    text: str
    box: Box
    confidence: float
    line_key: Tuple[int, int, int, int]


@dataclass
class OCRLine:
    text: str
    words: List[OCRWord]
    box: Box


@dataclass
class InvoiceDateMatch:
    label_box: Box
    date_box: Box
    invoice_date: date
    raw_text: str
    score: float

    @property
    def due_date(self) -> date:
        return self.invoice_date + timedelta(days=30)


def default_due_date_keywords() -> List[str]:
    return ["invoice date", "inv date"]


@dataclass
class DueDateRuleConfig:
    enabled: bool = True
    name: str = "Due Date Rule"
    label_text: str = "Due Date"
    offset_days: int = 30
    detection_mode: str = "auto"
    label_keywords: List[str] = field(default_factory=default_due_date_keywords)
    font_family: str = "Arial"
    font_size_adjust: int = 2
    line_gap_adjust: int = 6
    x_offset: int = -30
    y_offset: int = 0
    mirrored_margin: bool = True


@dataclass
class NoteRuleConfig:
    enabled: bool = True
    name: str = "Bottom Note"
    text: str = NOTE_TEXT
    font_family: str = "Arial"
    font_size_adjust: int = 5
    centered: bool = True
    use_body_margins: bool = True
    bottom_margin_adjust: int = 0


@dataclass
class ProcessingConfig:
    due_date_rule: Optional[DueDateRuleConfig] = field(default_factory=DueDateRuleConfig)
    note_rule: Optional[NoteRuleConfig] = field(default_factory=NoteRuleConfig)


RuleConfig = DueDateRuleConfig | NoteRuleConfig


def default_rule_configs() -> List[RuleConfig]:
    return [DueDateRuleConfig(), NoteRuleConfig()]


def default_processing_config() -> ProcessingConfig:
    return ProcessingConfig()


def processing_config_from_rules(rules: Sequence[RuleConfig]) -> ProcessingConfig:
    due_date_rule = next((rule for rule in rules if isinstance(rule, DueDateRuleConfig)), None)
    note_rule = next((rule for rule in rules if isinstance(rule, NoteRuleConfig)), None)
    return ProcessingConfig(
        due_date_rule=due_date_rule,
        note_rule=note_rule,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Add OCR-based due dates and a fixed payment note to scanned invoice PDFs."
    )
    parser.add_argument(
        "--input",
        type=Path,
        help="Path to a single input PDF. If omitted, all PDFs in INPUT are processed.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Path for a single output PDF. Only valid with --input.",
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=DEFAULT_INPUT_DIR,
        help="Directory to scan for PDFs when --input is not provided. Default: INPUT",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help="Directory for generated PDFs when --output is not provided. Default: OUTPUT",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=300,
        help="Render DPI used for OCR and output quality. Default: 300",
    )
    parser.add_argument(
        "--jpeg-quality",
        type=int,
        default=95,
        help="JPEG quality used when rebuilding annotated pages. Default: 95",
    )
    parser.add_argument(
        "--tesseract-cmd",
        help="Explicit path to the Tesseract executable if it is not on PATH.",
    )
    parser.add_argument(
        "--note-every-page",
        action="store_true",
        default=True,
        help="Always add the fixed payment note to every page. Enabled by default.",
    )
    return parser.parse_args()


def normalize_token(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", text.lower())


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", text.lower())).strip()


def union_boxes(boxes: Iterable[Box]) -> Box:
    boxes = list(boxes)
    if not boxes:
        raise ValueError("Cannot union an empty set of boxes.")
    return Box(
        min(box.x0 for box in boxes),
        min(box.y0 for box in boxes),
        max(box.x1 for box in boxes),
        max(box.y1 for box in boxes),
    )


def configure_tesseract(tesseract_cmd: Optional[str]) -> None:
    if tesseract_cmd:
        resolved = shutil.which(tesseract_cmd)
        path = Path(resolved) if resolved else Path(tesseract_cmd)
        if not path.exists():
            raise FileNotFoundError("Tesseract executable not found: %s" % tesseract_cmd)
        pytesseract.pytesseract.tesseract_cmd = str(path)
        return

    if shutil.which("tesseract"):
        return

    raise FileNotFoundError(
        "Tesseract OCR was not found on PATH. Install Tesseract or pass --tesseract-cmd."
    )


def resolve_font_path(preferred_family: str, bold: bool = False) -> Optional[Path]:
    if not preferred_family:
        return None

    fonts_dir = Path(r"C:\Windows\Fonts")
    if not fonts_dir.exists():
        return None

    preferred = re.sub(r"[^a-z0-9]+", "", preferred_family.lower())
    if any(keyword in preferred for keyword in UNSAFE_FONT_KEYWORDS):
        return None

    candidates = []
    for path in fonts_dir.glob("*.ttf"):
        stem = re.sub(r"[^a-z0-9]+", "", path.stem.lower())
        if preferred and (preferred in stem or stem in preferred):
            if any(keyword in stem for keyword in UNSAFE_FONT_KEYWORDS):
                continue
            score = 0
            if stem == preferred:
                score += 100
            elif stem.startswith(preferred):
                score += 50
            elif preferred.startswith(stem):
                score += 20
            if bold and ("bold" in path.stem.lower() or "bd" in path.stem.lower()):
                score += 10
            if not bold and "bold" not in path.stem.lower():
                score += 5
            candidates.append((score, -len(path.name), path))

    if not candidates:
        return None

    candidates.sort(reverse=True)
    return candidates[0][2]


def load_font(size: int, bold: bool = False, preferred_family: str = "") -> ImageFont.ImageFont:
    size = max(12, int(size))
    preferred_path = resolve_font_path(preferred_family, bold=bold)
    if preferred_path is not None:
        return ImageFont.truetype(str(preferred_path), size=size)

    candidates = FONT_PATHS["bold" if bold else "regular"]
    for path in candidates:
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    try:
        fallback_name = "arialbd.ttf" if bold else "arial.ttf"
        return ImageFont.truetype(fallback_name, size=size)
    except OSError:
        return ImageFont.load_default()


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_width: int) -> List[str]:
    words = text.split()
    if not words:
        return []

    lines = []
    current = words[0]

    for word in words[1:]:
        trial = "%s %s" % (current, word)
        if draw.textlength(trial, font=font) <= max_width:
            current = trial
        else:
            lines.append(current)
            current = word

    lines.append(current)
    return lines


def line_height(draw: ImageDraw.ImageDraw, font: ImageFont.ImageFont) -> int:
    _, top, _, bottom = draw.textbbox((0, 0), "Ag", font=font)
    return max(1, bottom - top)


def preprocess_for_ocr(image: Image.Image) -> Image.Image:
    grayscale = ImageOps.grayscale(image)
    grayscale = ImageOps.autocontrast(grayscale)
    return grayscale


def render_page_to_image(page: fitz.Page, dpi: int) -> Image.Image:
    pix = page.get_pixmap(dpi=dpi, alpha=False)
    mode = "RGB" if pix.n >= 3 else "L"
    image = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
    if image.mode != "RGB":
        image = image.convert("RGB")
    return image


def extract_ocr_lines(image: Image.Image, config: str) -> List[OCRLine]:
    data = pytesseract.image_to_data(image, output_type=Output.DICT, config=config)
    grouped = {}
    total = len(data["text"])

    for index in range(total):
        text = (data["text"][index] or "").strip()
        if not text:
            continue

        try:
            confidence = float(data["conf"][index])
        except (TypeError, ValueError):
            confidence = -1.0

        if confidence < 0:
            continue

        word = OCRWord(
            text=text,
            box=Box(
                int(data["left"][index]),
                int(data["top"][index]),
                int(data["left"][index]) + int(data["width"][index]),
                int(data["top"][index]) + int(data["height"][index]),
            ),
            confidence=confidence,
            line_key=(
                int(data["page_num"][index]),
                int(data["block_num"][index]),
                int(data["par_num"][index]),
                int(data["line_num"][index]),
            ),
        )
        grouped.setdefault(word.line_key, []).append(word)

    lines = []
    for words in grouped.values():
        words.sort(key=lambda item: (item.box.x0, item.box.y0))
        lines.append(
            OCRLine(
                text=" ".join(word.text for word in words),
                words=words,
                box=union_boxes(word.box for word in words),
            )
        )

    lines.sort(key=lambda line: (line.box.y0, line.box.x0))
    return lines


def invoice_label_score(text: str, label_keywords: Optional[Sequence[str]] = None) -> float:
    normalized = normalize_text(text)
    compact = normalize_token(text)
    keywords = list(label_keywords or default_due_date_keywords())

    best_score = 0.0
    for keyword in keywords:
        keyword_normalized = normalize_text(keyword)
        keyword_compact = normalize_token(keyword)
        if keyword_compact and keyword_compact in compact:
            best_score = max(best_score, 110.0 if len(keyword_compact) > 6 else 100.0)
        if keyword_normalized and keyword_normalized in normalized:
            best_score = max(best_score, 100.0)

    if best_score > 0:
        return best_score

    if "invoicedate" in compact:
        return 110.0
    if re.search(r"\binvoice\s+date\b", normalized):
        return 100.0
    if re.search(r"\binv\s+date\b", normalized):
        return 95.0
    return 0.0


def looks_like_date_word(text: str) -> bool:
    cleaned = text.strip().strip(",.;:")
    normalized = cleaned.lower()
    if not cleaned:
        return False
    if re.search(r"\d{1,4}", cleaned.translate(NUMERIC_DATE_FIXUPS)) and len(re.sub(r"\D", "", cleaned)) <= 4:
        return True
    if re.search(r"\d[\/\-.]\d", cleaned.translate(NUMERIC_DATE_FIXUPS)):
        return True
    return bool(re.fullmatch(MONTH_PATTERN, normalized))


def first_date_from_words(words: Sequence[OCRWord]) -> Optional[Tuple[str, date, Box]]:
    for start_index in range(len(words)):
        if not looks_like_date_word(words[start_index].text):
            continue

        for end_index in range(start_index, min(len(words), start_index + 4)):
            if end_index > start_index and not looks_like_date_word(words[end_index].text):
                break

            collected = words[start_index : end_index + 1]
            candidate_text = " ".join(word.text for word in collected)
            parsed = try_parse_date(candidate_text)
            if parsed:
                return candidate_text, parsed, union_boxes(word.box for word in collected)

    joined_text = " ".join(word.text for word in words)
    candidates = extract_date_candidates(joined_text)
    if candidates:
        raw_text, parsed = candidates[0]
        return raw_text, parsed, union_boxes(word.box for word in words)

    return None


def extract_date_candidates(text: str) -> List[Tuple[str, date]]:
    candidates = []
    seen = set()

    for pattern in DATE_PATTERNS:
        for match in pattern.finditer(text):
            raw_value = match.group(0).strip()
            parsed = try_parse_date(raw_value)
            if parsed and (raw_value, parsed) not in seen:
                seen.add((raw_value, parsed))
                candidates.append((raw_value, parsed))

    return candidates


def try_parse_date(raw_value: str) -> Optional[date]:
    cleaned = re.sub(r"\s+", " ", raw_value.strip())
    without_commas = cleaned.replace(",", "")
    normalized_numeric = without_commas.translate(NUMERIC_DATE_FIXUPS)
    probes = [
        without_commas,
        without_commas.title(),
        normalized_numeric,
        normalized_numeric.title(),
    ]

    for probe in probes:
        for date_format in NUMERIC_DATE_FORMATS:
            try:
                return datetime.strptime(probe, date_format).date()
            except ValueError:
                pass

        for date_format in ALPHA_DATE_FORMATS:
            try:
                return datetime.strptime(probe, date_format).date()
            except ValueError:
                pass

    return None


def format_due_date_like_source(source_text: str, due_date: date) -> str:
    source_text = source_text.strip()
    numeric_source = source_text.translate(NUMERIC_DATE_FIXUPS)

    if re.fullmatch(r"\d{4}[\/\-.]\d{1,2}[\/\-.]\d{1,2}", numeric_source):
        separator = re.search(r"[\/\-.]", numeric_source).group(0)
        return due_date.strftime("%%Y%s%%m%s%%d" % (separator, separator))

    if re.fullmatch(r"\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}", numeric_source):
        separator = re.search(r"[\/\-.]", numeric_source).group(0)
        year_token = re.split(r"[\/\-.]", numeric_source)[2]
        year_format = "%y" if len(year_token) == 2 else "%Y"
        return due_date.strftime("%%m%s%%d%s%s" % (separator, separator, year_format))

    lower = source_text.lower()
    has_comma = "," in source_text
    is_abbreviated_month = bool(
        re.match(r"\b(?:jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\b", lower)
    )
    month_format = "%b" if is_abbreviated_month else "%B"

    if re.match(r"^\d{1,2}\s+", lower):
        return due_date.strftime("%%d %s %%Y" % month_format)

    return due_date.strftime("%s %%d%s %%Y" % (month_format, "," if has_comma else ""))


def find_invoice_date_match(
    lines: Sequence[OCRLine],
    image_width: int,
    image_height: int,
    due_date_rule: Optional[DueDateRuleConfig] = None,
) -> Optional[InvoiceDateMatch]:
    detection_mode = (due_date_rule.detection_mode if due_date_rule else "auto").lower()
    label_keywords = due_date_rule.label_keywords if due_date_rule else default_due_date_keywords()
    date_lines = []
    for index, line in enumerate(lines):
        if extract_date_candidates(line.text):
            date_lines.append((index, line))

    best_match = None

    if detection_mode in {"auto", "invoice_label"}:
        for index, line in enumerate(lines):
            score = invoice_label_score(line.text, label_keywords=label_keywords)
            if score <= 0:
                continue

            label_words = [
                word
                for word in line.words
                if normalize_token(word.text) in {"invoice", "inv", "date", "invoicedate", "invdate"}
            ]
            label_box = union_boxes(word.box for word in label_words) if label_words else line.box

            same_line_words = [word for word in line.words if word.box.x0 >= label_box.x0]
            same_line_date = first_date_from_words(same_line_words)
            if same_line_date:
                raw_text, parsed_date, date_box = same_line_date
                candidate = InvoiceDateMatch(
                    label_box=label_box,
                    date_box=date_box,
                    invoice_date=parsed_date,
                    raw_text=raw_text,
                    score=score + 120.0,
                )
                if best_match is None or candidate.score > best_match.score:
                    best_match = candidate

            for candidate_index, candidate_line in date_lines:
                if candidate_index == index:
                    continue

                vertical_gap = candidate_line.box.y0 - label_box.y1
                if vertical_gap < -label_box.height:
                    continue

                if vertical_gap > max(label_box.height * 5, int(image_height * 0.08)):
                    continue

                horizontal_gap = abs(candidate_line.box.x0 - label_box.x0)
                if horizontal_gap > int(image_width * 0.25):
                    continue

                candidate_date = first_date_from_words(candidate_line.words)
                if not candidate_date:
                    continue

                raw_text, parsed_date, date_box = candidate_date
                candidate_score = score + 90.0 - (vertical_gap * 0.2) - (horizontal_gap * 0.05)
                candidate = InvoiceDateMatch(
                    label_box=label_box,
                    date_box=date_box,
                    invoice_date=parsed_date,
                    raw_text=raw_text,
                    score=candidate_score,
                )
                if best_match is None or candidate.score > best_match.score:
                    best_match = candidate

    if best_match is not None:
        return best_match

    if detection_mode == "invoice_label":
        return None

    page_has_invoice_context = any("invoice" in normalize_text(line.text) for line in lines)
    if not page_has_invoice_context:
        return None

    fallback_match = None
    top_band_limit = int(image_height * 0.25)

    for _, line in date_lines:
        if line.box.y1 > top_band_limit:
            continue

        candidate_date = first_date_from_words(line.words)
        if not candidate_date:
            continue

        raw_text, parsed_date, date_box = candidate_date
        normalized_line = normalize_text(line.text)
        normalized_raw = normalize_text(raw_text)
        exact_line_bonus = 20.0 if normalized_line == normalized_raw else 8.0
        candidate_score = 60.0 + exact_line_bonus - (line.box.y0 * 0.05) - (line.box.x0 * 0.01)

        candidate = InvoiceDateMatch(
            label_box=date_box,
            date_box=date_box,
            invoice_date=parsed_date,
            raw_text=raw_text,
            score=candidate_score,
        )
        if fallback_match is None or candidate.score > fallback_match.score:
            fallback_match = candidate

    return fallback_match


def fit_due_date_fonts(
    draw: ImageDraw.ImageDraw,
    label_text: str,
    value_text: str,
    max_width: int,
    base_size: int,
    preferred_family: str = "",
) -> Tuple[ImageFont.ImageFont, ImageFont.ImageFont, Tuple[int, int, int, int], Tuple[int, int, int, int]]:
    for size in range(base_size, 11, -2):
        label_font = load_font(max(12, size - 4), preferred_family=preferred_family)
        value_font = load_font(size, preferred_family=preferred_family)
        label_bbox = draw.textbbox((0, 0), label_text, font=label_font)
        value_bbox = draw.textbbox((0, 0), value_text, font=value_font)
        block_width = max(label_bbox[2] - label_bbox[0], value_bbox[2] - value_bbox[0])
        if block_width <= max_width:
            return label_font, value_font, label_bbox, value_bbox

    label_font = load_font(12, preferred_family=preferred_family)
    value_font = load_font(12, preferred_family=preferred_family)
    return (
        label_font,
        value_font,
        draw.textbbox((0, 0), label_text, font=label_font),
        draw.textbbox((0, 0), value_text, font=value_font),
    )


def median_int(values: Sequence[int], fallback: int) -> int:
    if not values:
        return fallback
    ordered = sorted(values)
    middle = len(ordered) // 2
    if len(ordered) % 2 == 1:
        return ordered[middle]
    return int(round((ordered[middle - 1] + ordered[middle]) / 2.0))


def estimate_document_font_size(lines: Sequence[OCRLine], image_height: int) -> int:
    body_heights = [
        line.box.height
        for line in lines
        if 2 <= len(line.words) <= 20
        and line.box.y0 >= int(image_height * 0.08)
        and line.box.y1 <= int(image_height * 0.85)
    ]
    if not body_heights:
        body_heights = [line.box.height for line in lines if line.box.height > 0]

    median_height = median_int(body_heights, fallback=36)
    return max(16, min(36, int(round(median_height * 0.78))))


def draw_due_date_block(
    image: Image.Image,
    match: InvoiceDateMatch,
    document_font_size: int,
    due_date_rule: DueDateRuleConfig,
) -> None:
    draw = ImageDraw.Draw(image)
    image_width, image_height = image.size
    anchor_box = union_boxes([match.label_box, match.date_box])
    due_date_value = match.invoice_date + timedelta(days=due_date_rule.offset_days)
    due_date_text = format_due_date_like_source(match.raw_text, due_date_value)

    label_text = due_date_rule.label_text or "Due Date"
    value_text = due_date_text
    margin = max(24, image_width // 60)
    gap = max(40, image_width // 40)
    _ = document_font_size
    base_font_size = (
        max(24, min(56, int(max(anchor_box.height, match.date_box.height) * 0.9)))
        + due_date_rule.font_size_adjust
    )
    mirrored_right_inset = max(margin, anchor_box.x0) if due_date_rule.mirrored_margin else margin
    available_width = max(120, image_width - margin - mirrored_right_inset)

    label_font, value_font, label_bbox, value_bbox = fit_due_date_fonts(
        draw=draw,
        label_text=label_text,
        value_text=value_text,
        max_width=available_width,
        base_size=base_font_size,
        preferred_family=due_date_rule.font_family,
    )

    label_width = label_bbox[2] - label_bbox[0]
    label_height = label_bbox[3] - label_bbox[1]
    value_width = value_bbox[2] - value_bbox[0]
    value_height = value_bbox[3] - value_bbox[1]
    block_width = max(label_width, value_width)
    resolved_font_size = max(12, getattr(value_font, "size", base_font_size))
    line_gap = max(4, (resolved_font_size // 4) + due_date_rule.line_gap_adjust)
    padding_x = max(10, resolved_font_size // 2)
    padding_y = max(4, resolved_font_size // 4)
    block_height = label_height + line_gap + value_height

    x = image_width - mirrored_right_inset - block_width - padding_x
    minimum_x = anchor_box.x1 + gap + padding_x
    maximum_x = image_width - mirrored_right_inset - block_width - padding_x
    x = max(margin, min(maximum_x, max(x, minimum_x)))
    x += due_date_rule.x_offset
    y = max(margin, anchor_box.y0 - max(0, (block_height - anchor_box.height) // 2))
    y += due_date_rule.y_offset
    if y + block_height + (padding_y * 2) > image_height - margin:
        y = max(margin, image_height - margin - block_height - (padding_y * 2))
    x = max(margin, min(image_width - block_width - padding_x, x))
    y = max(margin, min(image_height - block_height - padding_y, y))

    draw.text((x, y), label_text, fill="black", font=label_font)
    draw.text((x, y + label_height + line_gap), value_text, fill="black", font=value_font)


def estimate_note_margins(
    lines: Sequence[OCRLine],
    image_width: int,
    image_height: int,
) -> Tuple[int, int]:
    candidate_lines = [
        line
        for line in lines
        if len(line.words) >= 3
        and line.box.y0 >= int(image_height * 0.12)
        and line.box.y1 <= int(image_height * 0.82)
    ]

    left_candidates = [line.box.x0 for line in candidate_lines if line.box.x0 < int(image_width * 0.35)]
    right_candidates = [
        image_width - line.box.x1
        for line in candidate_lines
        if line.box.x1 > int(image_width * 0.65)
    ]

    fallback_margin = max(32, image_width // 24)
    left_margin = median_int(left_candidates, fallback=fallback_margin)
    right_margin = median_int(right_candidates, fallback=left_margin)
    return left_margin, right_margin


def draw_note_block(
    image: Image.Image,
    document_font_size: int,
    lines: Sequence[OCRLine],
    note_rule: NoteRuleConfig,
) -> None:
    draw = ImageDraw.Draw(image)
    image_width, image_height = image.size
    margin_bottom = max(28, image_height // 28) + note_rule.bottom_margin_adjust
    if note_rule.use_body_margins:
        left_margin, right_margin = estimate_note_margins(lines, image_width, image_height)
    else:
        left_margin = right_margin = max(32, image_width // 24)
    max_width = image_width - left_margin - right_margin
    _ = document_font_size
    font_size = max(18, min(36, image_width // 55)) + note_rule.font_size_adjust
    font = load_font(font_size, preferred_family=note_rule.font_family)
    lines = wrap_text(draw, note_rule.text, font, max_width)
    text_height = line_height(draw, font)
    line_gap = max(4, font_size // 4)
    text_block_height = (len(lines) * text_height) + (max(0, len(lines) - 1) * line_gap)
    top = max(10, image_height - margin_bottom - text_block_height)

    y = top
    for line in lines:
        line_width = draw.textlength(line, font=font)
        if note_rule.centered:
            x = int(round(left_margin + ((max_width - line_width) / 2.0)))
        else:
            x = left_margin
        draw.text((x, y), line, fill="black", font=font)
        y += text_height + line_gap


def annotate_page(
    image: Image.Image,
    match: Optional[InvoiceDateMatch],
    lines: Sequence[OCRLine],
    add_note: bool,
    config: Optional[ProcessingConfig] = None,
) -> Image.Image:
    annotated = image.copy()
    config = config or default_processing_config()
    document_font_size = estimate_document_font_size(lines, image.height)
    if match is not None and config.due_date_rule and config.due_date_rule.enabled:
        draw_due_date_block(annotated, match, document_font_size, config.due_date_rule)
    if add_note and config.note_rule and config.note_rule.enabled:
        draw_note_block(annotated, document_font_size, lines, config.note_rule)
    return annotated


def pil_image_to_jpeg_bytes(image: Image.Image, quality: int) -> bytes:
    buffer = io.BytesIO()
    image.save(buffer, format="JPEG", quality=quality, optimize=True, subsampling=0)
    return buffer.getvalue()


def resolve_work_items(args: argparse.Namespace) -> List[Tuple[Path, Path]]:
    if args.input and not args.output:
        return [
            (
                args.input,
                args.output_dir / ("%s_due_date.pdf" % args.input.stem),
            )
        ]

    if args.input and args.output:
        return [(args.input, args.output)]

    if args.output and not args.input:
        raise ValueError("--output can only be used together with --input.")

    input_dir = args.input_dir
    output_dir = args.output_dir
    pdfs = sorted(input_dir.glob("*.pdf"))
    return [(pdf, output_dir / ("%s_due_date.pdf" % pdf.stem)) for pdf in pdfs]


def process_pdf(
    input_pdf: Path,
    output_pdf: Path,
    dpi: int,
    jpeg_quality: int,
    add_note: bool,
    config: Optional[ProcessingConfig] = None,
    log_callback: Optional[Callable[[str], None]] = None,
) -> None:
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    processing_config = config or default_processing_config()

    def emit_status(message: str) -> None:
        if log_callback is not None:
            log_callback(message)
        else:
            sys.stdout.write(message + "\n")

    source = fitz.open(str(input_pdf))
    target = fitz.open()

    try:
        for page_index in range(source.page_count):
            page = source.load_page(page_index)
            image = render_page_to_image(page, dpi=dpi)
            ocr_image = preprocess_for_ocr(image)

            lines = []
            match = None
            for ocr_config in OCR_CONFIGS:
                lines = extract_ocr_lines(ocr_image, ocr_config)
                match = find_invoice_date_match(
                    lines,
                    image.width,
                    image.height,
                    due_date_rule=processing_config.due_date_rule,
                )
                if match is not None:
                    break

            annotated = annotate_page(image, match, lines, add_note=add_note, config=processing_config)
            image_bytes = pil_image_to_jpeg_bytes(annotated, quality=jpeg_quality)

            new_page = target.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(new_page.rect, stream=image_bytes, keep_proportion=False)

            if match is not None and processing_config.due_date_rule and processing_config.due_date_rule.enabled:
                due_date_value = match.invoice_date + timedelta(days=processing_config.due_date_rule.offset_days)
                emit_status(
                    "[OK] %s - page %d: invoice date %s -> due date %s"
                    % (
                        input_pdf.name,
                        page_index + 1,
                        match.invoice_date.isoformat(),
                        due_date_value.isoformat(),
                    )
                )
            else:
                emit_status(
                    "[WARN] %s - page %d: invoice date not found, note added only"
                    % (input_pdf.name, page_index + 1)
                )

        target.save(str(output_pdf), deflate=True, garbage=4)
    finally:
        source.close()
        target.close()


def main() -> int:
    args = parse_args()

    try:
        configure_tesseract(args.tesseract_cmd)
        work_items = resolve_work_items(args)
    except Exception as exc:
        sys.stderr.write("Error: %s\n" % exc)
        return 1

    if not work_items:
        sys.stderr.write("No PDF files were found to process.\n")
        return 1

    for input_pdf, output_pdf in work_items:
        if not input_pdf.exists():
            sys.stderr.write("Skipping missing file: %s\n" % input_pdf)
            continue
        try:
            process_pdf(
                input_pdf=input_pdf,
                output_pdf=output_pdf,
                dpi=args.dpi,
                jpeg_quality=args.jpeg_quality,
                add_note=args.note_every_page,
            )
            sys.stdout.write("Saved: %s\n" % output_pdf)
        except Exception as exc:
            sys.stderr.write("Failed to process %s: %s\n" % (input_pdf, exc))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
