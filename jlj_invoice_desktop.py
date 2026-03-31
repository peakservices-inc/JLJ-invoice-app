#!/usr/bin/env python3
from __future__ import annotations

import json
import os
import sys
import traceback
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional, Sequence

from PySide6 import QtCore, QtGui, QtWidgets

import annotate_invoice_due_dates as backend


APP_TITLE = "JLJ IV Enterprises Inc. Invoice Rule Studio"
APP_SUBTITLE = "Property of JLJ IV Enterprises Inc."
APP_DIR = Path(os.environ.get("APPDATA", str(Path.home()))).joinpath(
    "JLJ IV Enterprises Inc", "Invoice Rule Studio"
)
STATE_PATH = APP_DIR / "settings.json"
LOG_PATH = APP_DIR / "app.log"
RuleObject = backend.RuleConfig


def ensure_app_dir() -> None:
    APP_DIR.mkdir(parents=True, exist_ok=True)


def append_app_log(message: str) -> None:
    ensure_app_dir()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with LOG_PATH.open("a", encoding="utf-8") as handle:
        handle.write(f"[{timestamp}] {message.rstrip()}\n")


def default_tesseract_path() -> str:
    candidates = [
        Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
        Path(r"C:\Users\%s\AppData\Local\Programs\Tesseract-OCR\tesseract.exe" % os.environ.get("USERNAME", "")),
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    return ""


def normalize_preview(text: str, max_chars: int = 88) -> str:
    compact = " ".join(text.split())
    if len(compact) <= max_chars:
        return compact
    return compact[: max_chars - 3].rstrip() + "..."


def configure_form_layout(form: QtWidgets.QFormLayout) -> None:
    form.setLabelAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
    form.setFormAlignment(QtCore.Qt.AlignTop)
    form.setSpacing(10)
    form.setFieldGrowthPolicy(QtWidgets.QFormLayout.ExpandingFieldsGrow)
    form.setRowWrapPolicy(QtWidgets.QFormLayout.WrapLongRows)


def create_scroll_form_page() -> tuple[QtWidgets.QScrollArea, QtWidgets.QWidget, QtWidgets.QFormLayout]:
    content = QtWidgets.QWidget()
    form = QtWidgets.QFormLayout(content)
    configure_form_layout(form)

    scroll = QtWidgets.QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
    scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
    scroll.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
    scroll.setWidget(content)
    return scroll, content, form


def set_compact_input_height(widget: QtWidgets.QWidget, min_height: int = 32) -> None:
    widget.setMinimumHeight(min_height)


def safe_font_family_name(family: str, fallback: str = "Arial") -> str:
    normalized = "".join(ch for ch in family.lower() if ch.isalnum())
    if not normalized:
        return fallback
    if any(keyword in normalized for keyword in backend.UNSAFE_FONT_KEYWORDS):
        return fallback
    return family


def sanitize_rule(rule: RuleObject) -> RuleObject:
    if isinstance(rule, backend.DueDateRuleConfig):
        rule.font_family = safe_font_family_name(rule.font_family or "Arial")
        return rule
    if isinstance(rule, backend.NoteRuleConfig):
        rule.font_family = safe_font_family_name(rule.font_family or "Arial")
        return rule
    return rule


class ProcessWorker(QtCore.QObject):
    finished = QtCore.Signal(list)
    failed = QtCore.Signal(str)
    log = QtCore.Signal(str)
    progress = QtCore.Signal(int, int)

    def __init__(
        self,
        input_files: List[str],
        output_dir: str,
        dpi: int,
        jpeg_quality: int,
        tesseract_path: str,
        rules: Sequence[RuleObject],
    ) -> None:
        super().__init__()
        self.input_files = input_files
        self.output_dir = Path(output_dir)
        self.dpi = dpi
        self.jpeg_quality = jpeg_quality
        self.tesseract_path = tesseract_path
        self.rules = list(rules)

    @QtCore.Slot()
    def run(self) -> None:
        outputs = []
        try:
            if self.tesseract_path:
                backend.configure_tesseract(self.tesseract_path)

            config = backend.processing_config_from_rules(self.rules)

            total = len(self.input_files)
            self.output_dir.mkdir(parents=True, exist_ok=True)
            add_note = any(
                isinstance(rule, backend.NoteRuleConfig) and rule.enabled
                for rule in self.rules
            )

            for index, input_file in enumerate(self.input_files, start=1):
                input_path = Path(input_file)
                output_path = self.output_dir / f"{input_path.stem}_processed.pdf"
                self.log.emit(f"Processing {input_path.name}...")
                backend.process_pdf(
                    input_pdf=input_path,
                    output_pdf=output_path,
                    dpi=self.dpi,
                    jpeg_quality=self.jpeg_quality,
                    add_note=add_note,
                    config=config,
                    log_callback=self.log.emit,
                )
                outputs.append(str(output_path))
                self.progress.emit(index, total)
                self.log.emit(f"Finished {input_path.name}")

            self.finished.emit(outputs)
        except Exception:
            self.failed.emit(traceback.format_exc())


class CardFrame(QtWidgets.QFrame):
    def __init__(self, title: str, subtitle: str = "") -> None:
        super().__init__()
        self.setObjectName("Card")
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        title_label = QtWidgets.QLabel(title)
        title_label.setObjectName("CardTitle")
        layout.addWidget(title_label)

        if subtitle:
            subtitle_label = QtWidgets.QLabel(subtitle)
            subtitle_label.setObjectName("CardSubtitle")
            subtitle_label.setWordWrap(True)
            layout.addWidget(subtitle_label)

        self.body_layout = layout


class DueDateRuleEditor(QtWidgets.QWidget):
    changed = QtCore.Signal()

    def __init__(self) -> None:
        super().__init__()
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        intro = QtWidgets.QLabel(
            "Use this rule to find the invoice date and place a due date beside it."
        )
        intro.setWordWrap(True)
        intro.setObjectName("CardSubtitle")
        layout.addWidget(intro)

        self.tabs = QtWidgets.QTabWidget()
        self.tabs.setDocumentMode(True)
        layout.addWidget(self.tabs)

        basic_scroll, basic_page, basic_form = create_scroll_form_page()

        self.enabled = QtWidgets.QCheckBox("Turn this rule on")
        self.name_edit = QtWidgets.QLineEdit()
        self.label_text = QtWidgets.QLineEdit()
        self.offset_days = QtWidgets.QSpinBox()
        self.offset_days.setRange(1, 365)
        self.offset_days.setValue(30)

        self.detection_mode = QtWidgets.QComboBox()
        self.detection_mode.addItem("Auto detect", "auto")
        self.detection_mode.addItem("Month name at top (example: March 19, 2026)", "top_date")
        self.detection_mode.addItem("Invoice Date label", "invoice_label")
        self.detection_mode.setSizeAdjustPolicy(QtWidgets.QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.detection_mode.setMinimumContentsLength(20)

        self.label_keywords = QtWidgets.QLineEdit()
        self.label_keywords.setPlaceholderText("invoice date, inv date")

        help_label = QtWidgets.QLabel(
            "Tip: use commas between label names. Example: invoice date, inv date"
        )
        help_label.setWordWrap(True)
        help_label.setObjectName("CardSubtitle")

        basic_form.addRow("", self.enabled)
        basic_form.addRow("Rule name", self.name_edit)
        basic_form.addRow("Text to place", self.label_text)
        basic_form.addRow("Days after invoice date", self.offset_days)
        basic_form.addRow("How to find the invoice date", self.detection_mode)
        basic_form.addRow("Other date labels to look for", self.label_keywords)
        basic_form.addRow("", help_label)

        advanced_scroll, advanced_page, advanced_form = create_scroll_form_page()

        self.font_family = QtWidgets.QFontComboBox()
        self.font_size_adjust = QtWidgets.QSpinBox()
        self.font_size_adjust.setRange(-12, 24)
        self.line_gap_adjust = QtWidgets.QSpinBox()
        self.line_gap_adjust.setRange(-20, 20)
        self.x_offset = QtWidgets.QSpinBox()
        self.x_offset.setRange(-500, 500)
        self.y_offset = QtWidgets.QSpinBox()
        self.y_offset.setRange(-500, 500)
        self.mirrored_margin = QtWidgets.QCheckBox("Match right-side spacing to the left-side date spacing")

        advanced_form.addRow("Font family", self.font_family)
        advanced_form.addRow("Text size adjustment", self.font_size_adjust)
        advanced_form.addRow("Space between label and value", self.line_gap_adjust)
        advanced_form.addRow("Move left or right", self.x_offset)
        advanced_form.addRow("Move up or down", self.y_offset)
        advanced_form.addRow("", self.mirrored_margin)

        for widget in [
            self.name_edit,
            self.label_text,
            self.offset_days,
            self.detection_mode,
            self.label_keywords,
            self.font_family,
            self.font_size_adjust,
            self.line_gap_adjust,
            self.x_offset,
            self.y_offset,
        ]:
            set_compact_input_height(widget)

        self.tabs.addTab(basic_scroll, "Basic")
        self.tabs.addTab(advanced_scroll, "Advanced")

        for widget in [
            self.enabled,
            self.name_edit,
            self.label_text,
            self.offset_days,
            self.detection_mode,
            self.label_keywords,
            self.font_family,
            self.font_size_adjust,
            self.line_gap_adjust,
            self.x_offset,
            self.y_offset,
            self.mirrored_margin,
        ]:
            if isinstance(widget, QtWidgets.QAbstractButton):
                widget.clicked.connect(lambda *_: self.changed.emit())
            elif isinstance(widget, QtWidgets.QLineEdit):
                widget.textChanged.connect(lambda *_: self.changed.emit())
            elif isinstance(widget, QtWidgets.QComboBox):
                widget.currentIndexChanged.connect(lambda *_: self.changed.emit())
            elif isinstance(widget, QtWidgets.QSpinBox):
                widget.valueChanged.connect(lambda *_: self.changed.emit())
            elif isinstance(widget, QtWidgets.QFontComboBox):
                widget.currentFontChanged.connect(lambda *_: self.changed.emit())

        self._rule: Optional[backend.DueDateRuleConfig] = None

    def set_rule(self, rule: backend.DueDateRuleConfig) -> None:
        self._rule = rule
        blockers = [
            QtCore.QSignalBlocker(self.enabled),
            QtCore.QSignalBlocker(self.name_edit),
            QtCore.QSignalBlocker(self.label_text),
            QtCore.QSignalBlocker(self.offset_days),
            QtCore.QSignalBlocker(self.detection_mode),
            QtCore.QSignalBlocker(self.label_keywords),
            QtCore.QSignalBlocker(self.font_family),
            QtCore.QSignalBlocker(self.font_size_adjust),
            QtCore.QSignalBlocker(self.line_gap_adjust),
            QtCore.QSignalBlocker(self.x_offset),
            QtCore.QSignalBlocker(self.y_offset),
            QtCore.QSignalBlocker(self.mirrored_margin),
        ]
        self.enabled.setChecked(rule.enabled)
        self.name_edit.setText(rule.name)
        self.label_text.setText(rule.label_text)
        self.offset_days.setValue(rule.offset_days)
        self.detection_mode.setCurrentIndex(max(0, self.detection_mode.findData(rule.detection_mode)))
        self.label_keywords.setText(", ".join(rule.label_keywords))
        self.font_family.setCurrentFont(QtGui.QFont(safe_font_family_name(rule.font_family or "Arial")))
        self.font_size_adjust.setValue(rule.font_size_adjust)
        self.line_gap_adjust.setValue(rule.line_gap_adjust)
        self.x_offset.setValue(rule.x_offset)
        self.y_offset.setValue(rule.y_offset)
        self.mirrored_margin.setChecked(rule.mirrored_margin)
        del blockers

    def apply_changes(self) -> None:
        if not self._rule:
            return
        self._rule.enabled = self.enabled.isChecked()
        self._rule.name = self.name_edit.text().strip() or "Due Date Rule"
        self._rule.label_text = self.label_text.text().strip() or "Due Date"
        self._rule.offset_days = self.offset_days.value()
        self._rule.detection_mode = self.detection_mode.currentData()
        self._rule.label_keywords = [
            item.strip()
            for item in self.label_keywords.text().split(",")
            if item.strip()
        ] or backend.default_due_date_keywords()
        self._rule.font_family = safe_font_family_name(self.font_family.currentFont().family())
        self._rule.font_size_adjust = self.font_size_adjust.value()
        self._rule.line_gap_adjust = self.line_gap_adjust.value()
        self._rule.x_offset = self.x_offset.value()
        self._rule.y_offset = self.y_offset.value()
        self._rule.mirrored_margin = self.mirrored_margin.isChecked()


class NoteRuleEditor(QtWidgets.QWidget):
    changed = QtCore.Signal()

    def __init__(self) -> None:
        super().__init__()
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        intro = QtWidgets.QLabel(
            "Use this rule to place a fixed note at the bottom of the page."
        )
        intro.setWordWrap(True)
        intro.setObjectName("CardSubtitle")
        layout.addWidget(intro)

        self.tabs = QtWidgets.QTabWidget()
        self.tabs.setDocumentMode(True)
        layout.addWidget(self.tabs)

        basic_scroll, basic_page, basic_form = create_scroll_form_page()

        self.enabled = QtWidgets.QCheckBox("Turn this rule on")
        self.name_edit = QtWidgets.QLineEdit()
        self.note_text = QtWidgets.QPlainTextEdit()
        self.note_text.setMinimumHeight(150)
        self.centered = QtWidgets.QCheckBox("Center the note inside the text area")

        basic_form.addRow("", self.enabled)
        basic_form.addRow("Rule name", self.name_edit)
        basic_form.addRow("Message to place", self.note_text)
        basic_form.addRow("", self.centered)

        advanced_scroll, advanced_page, advanced_form = create_scroll_form_page()

        self.font_family = QtWidgets.QFontComboBox()
        self.font_size_adjust = QtWidgets.QSpinBox()
        self.font_size_adjust.setRange(-12, 24)
        self.use_body_margins = QtWidgets.QCheckBox("Use the same left and right margins as the body text")
        self.bottom_margin_adjust = QtWidgets.QSpinBox()
        self.bottom_margin_adjust.setRange(-200, 200)

        advanced_form.addRow("Font family", self.font_family)
        advanced_form.addRow("Text size adjustment", self.font_size_adjust)
        advanced_form.addRow("", self.use_body_margins)
        advanced_form.addRow("Move higher or lower", self.bottom_margin_adjust)

        for widget in [
            self.name_edit,
            self.font_family,
            self.font_size_adjust,
            self.bottom_margin_adjust,
        ]:
            set_compact_input_height(widget)

        self.tabs.addTab(basic_scroll, "Basic")
        self.tabs.addTab(advanced_scroll, "Advanced")

        self.enabled.clicked.connect(lambda *_: self.changed.emit())
        self.name_edit.textChanged.connect(lambda *_: self.changed.emit())
        self.note_text.textChanged.connect(self.changed.emit)
        self.font_family.currentFontChanged.connect(lambda *_: self.changed.emit())
        self.font_size_adjust.valueChanged.connect(lambda *_: self.changed.emit())
        self.centered.clicked.connect(lambda *_: self.changed.emit())
        self.use_body_margins.clicked.connect(lambda *_: self.changed.emit())
        self.bottom_margin_adjust.valueChanged.connect(lambda *_: self.changed.emit())

        self._rule: Optional[backend.NoteRuleConfig] = None

    def set_rule(self, rule: backend.NoteRuleConfig) -> None:
        self._rule = rule
        blockers = [
            QtCore.QSignalBlocker(self.enabled),
            QtCore.QSignalBlocker(self.name_edit),
            QtCore.QSignalBlocker(self.note_text),
            QtCore.QSignalBlocker(self.font_family),
            QtCore.QSignalBlocker(self.font_size_adjust),
            QtCore.QSignalBlocker(self.centered),
            QtCore.QSignalBlocker(self.use_body_margins),
            QtCore.QSignalBlocker(self.bottom_margin_adjust),
        ]
        self.enabled.setChecked(rule.enabled)
        self.name_edit.setText(rule.name)
        self.note_text.setPlainText(rule.text)
        self.font_family.setCurrentFont(QtGui.QFont(safe_font_family_name(rule.font_family or "Arial")))
        self.font_size_adjust.setValue(rule.font_size_adjust)
        self.centered.setChecked(rule.centered)
        self.use_body_margins.setChecked(rule.use_body_margins)
        self.bottom_margin_adjust.setValue(rule.bottom_margin_adjust)
        del blockers

    def apply_changes(self) -> None:
        if not self._rule:
            return
        self._rule.enabled = self.enabled.isChecked()
        self._rule.name = self.name_edit.text().strip() or "Bottom Note"
        self._rule.text = self.note_text.toPlainText().strip() or backend.NOTE_TEXT
        self._rule.font_family = safe_font_family_name(self.font_family.currentFont().family())
        self._rule.font_size_adjust = self.font_size_adjust.value()
        self._rule.centered = self.centered.isChecked()
        self._rule.use_body_margins = self.use_body_margins.isChecked()
        self._rule.bottom_margin_adjust = self.bottom_margin_adjust.value()


@dataclass(frozen=True)
class RuleDefinition:
    rule_type: str
    display_name: str
    description: str
    config_type: type
    factory: Callable[[], RuleObject]
    editor_key: str
    max_instances: Optional[int] = 1


RULE_DEFINITIONS: List[RuleDefinition] = [
    RuleDefinition(
        rule_type="due_date",
        display_name="Due Date",
        description="Find the invoice date and place a due date beside it.",
        config_type=backend.DueDateRuleConfig,
        factory=backend.DueDateRuleConfig,
        editor_key="due_date",
        max_instances=1,
    ),
    RuleDefinition(
        rule_type="note",
        display_name="Bottom Note",
        description="Add a fixed note at the bottom of each page.",
        config_type=backend.NoteRuleConfig,
        factory=backend.NoteRuleConfig,
        editor_key="note",
        max_instances=1,
    ),
]

RULE_DEFINITION_BY_TYPE: Dict[str, RuleDefinition] = {
    definition.rule_type: definition for definition in RULE_DEFINITIONS
}


def starter_rules() -> List[RuleObject]:
    return backend.default_rule_configs()


def rule_definition_for_rule(rule: RuleObject) -> RuleDefinition:
    for definition in RULE_DEFINITIONS:
        if isinstance(rule, definition.config_type):
            return definition
    raise TypeError("Unsupported rule type")


def rule_to_dict(rule: RuleObject) -> dict:
    payload = asdict(rule)
    payload["rule_type"] = rule_definition_for_rule(rule).rule_type
    return payload


def rule_from_dict(payload: dict) -> RuleObject:
    rule_type = payload.get("rule_type")
    definition = RULE_DEFINITION_BY_TYPE.get(rule_type)
    if definition is None:
        raise ValueError("Unknown rule type: %s" % rule_type)

    clean = dict(payload)
    clean.pop("rule_type", None)
    return sanitize_rule(definition.config_type(**clean))


def clone_rules(rules: Sequence[RuleObject]) -> List[RuleObject]:
    return [rule_from_dict(rule_to_dict(rule)) for rule in rules]


def rule_summary(rule: RuleObject) -> str:
    if isinstance(rule, backend.DueDateRuleConfig):
        detection_labels = {
            "auto": "Auto detect",
            "top_date": "Top month-style date",
            "invoice_label": "Invoice Date label",
        }
        return (
            f"Adds '{rule.label_text}' {rule.offset_days} days after the invoice date. "
            f"Finding method: {detection_labels.get(rule.detection_mode, 'Auto detect')}."
        )

    if isinstance(rule, backend.NoteRuleConfig):
        position = "Centered at the bottom" if rule.centered else "Left aligned at the bottom"
        return f"{position}. Text: {normalize_preview(rule.text)}"

    return "Rule settings available."


def rule_list_text(rule: RuleObject) -> str:
    definition = rule_definition_for_rule(rule)
    status = "On" if getattr(rule, "enabled", True) else "Off"
    return f"{definition.display_name}: {rule.name} ({status})"


class AddRuleDialog(QtWidgets.QDialog):
    def __init__(self, definitions: Sequence[RuleDefinition], parent: Optional[QtWidgets.QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Add Rule")
        self.resize(420, 300)
        self._definitions = list(definitions)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        intro = QtWidgets.QLabel(
            "Choose the type of rule you want to add. New rule types will appear here automatically as the app grows."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        self.list_widget = QtWidgets.QListWidget()
        for definition in self._definitions:
            item = QtWidgets.QListWidgetItem(definition.display_name)
            item.setToolTip(definition.description)
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget, 1)

        self.description_label = QtWidgets.QLabel("Select a rule type to see what it does.")
        self.description_label.setWordWrap(True)
        self.description_label.setObjectName("CardSubtitle")
        layout.addWidget(self.description_label)

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        self.ok_button = buttons.button(QtWidgets.QDialogButtonBox.Ok)
        self.ok_button.setText("Add Rule")
        self.ok_button.setEnabled(False)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.list_widget.currentRowChanged.connect(self._update_selection)
        self.list_widget.itemDoubleClicked.connect(lambda *_: self.accept())

    def _update_selection(self, row: int) -> None:
        definition = self.selected_definition()
        self.ok_button.setEnabled(definition is not None)
        if definition is None:
            self.description_label.setText("Select a rule type to see what it does.")
            return
        self.description_label.setText(definition.description)

    def selected_definition(self) -> Optional[RuleDefinition]:
        row = self.list_widget.currentRow()
        if row < 0 or row >= len(self._definitions):
            return None
        return self._definitions[row]


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1480, 920)

        self.rules: List[RuleObject] = starter_rules()
        self.input_files: List[str] = []
        self.output_files: List[str] = []
        self.processing_thread: Optional[QtCore.QThread] = None
        self.worker: Optional[ProcessWorker] = None
        self.rule_editors: Dict[str, QtWidgets.QWidget] = {}

        self._build_ui()
        self._apply_styles()
        self._load_state()
        self._refresh_rule_list()
        self._refresh_input_list()
        self._refresh_output_list()

    def _build_ui(self) -> None:
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.setCentralWidget(scroll)

        root = QtWidgets.QWidget()
        root.setMinimumSize(1240, 860)
        scroll.setWidget(root)
        outer = QtWidgets.QVBoxLayout(root)
        outer.setContentsMargins(18, 18, 18, 18)
        outer.setSpacing(14)

        title = QtWidgets.QLabel(APP_TITLE)
        title.setObjectName("HeroTitle")
        subtitle = QtWidgets.QLabel(
            "A simpler PDF rule studio for office teams. Start with the starter rules, change only what you need, and keep advanced settings tucked away."
        )
        subtitle.setObjectName("HeroSubtitle")
        subtitle.setWordWrap(True)

        outer.addWidget(title)
        outer.addWidget(subtitle)

        splitter = QtWidgets.QSplitter()
        splitter.setOrientation(QtCore.Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        outer.addWidget(splitter, 1)

        left_column = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(left_column)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(14)
        splitter.addWidget(left_column)

        upload_card = CardFrame(
            "Step 1: Files",
            "Add one or more invoice PDFs to process.",
        )
        left_layout.addWidget(upload_card)

        self.input_list = QtWidgets.QListWidget()
        upload_card.body_layout.addWidget(self.input_list)

        upload_buttons = QtWidgets.QHBoxLayout()
        self.add_files_button = QtWidgets.QPushButton("Add Files")
        self.remove_file_button = QtWidgets.QPushButton("Remove Selected")
        self.clear_files_button = QtWidgets.QPushButton("Clear List")
        upload_buttons.addWidget(self.add_files_button)
        upload_buttons.addWidget(self.remove_file_button)
        upload_buttons.addWidget(self.clear_files_button)
        upload_card.body_layout.addLayout(upload_buttons)

        settings_card = CardFrame(
            "Step 2: Process",
            "Choose where finished PDFs should be saved. Most users can leave the OCR settings as they are.",
        )
        left_layout.addWidget(settings_card)

        settings_form = QtWidgets.QFormLayout()
        self.output_dir_edit = QtWidgets.QLineEdit(str(Path("OUTPUT").resolve()))
        self.output_dir_button = QtWidgets.QPushButton("Browse")
        output_dir_row = QtWidgets.QHBoxLayout()
        output_dir_row.addWidget(self.output_dir_edit)
        output_dir_row.addWidget(self.output_dir_button)
        output_dir_wrap = QtWidgets.QWidget()
        output_dir_wrap.setLayout(output_dir_row)

        self.tesseract_edit = QtWidgets.QLineEdit(default_tesseract_path())
        self.tesseract_button = QtWidgets.QPushButton("Browse")
        tesseract_row = QtWidgets.QHBoxLayout()
        tesseract_row.addWidget(self.tesseract_edit)
        tesseract_row.addWidget(self.tesseract_button)
        tesseract_wrap = QtWidgets.QWidget()
        tesseract_wrap.setLayout(tesseract_row)

        self.dpi_spin = QtWidgets.QSpinBox()
        self.dpi_spin.setRange(150, 600)
        self.dpi_spin.setValue(300)
        self.quality_spin = QtWidgets.QSpinBox()
        self.quality_spin.setRange(50, 100)
        self.quality_spin.setValue(95)

        settings_form.addRow("Save Processed Files To", output_dir_wrap)
        settings_form.addRow("OCR Program (Tesseract)", tesseract_wrap)
        settings_form.addRow("Render DPI", self.dpi_spin)
        settings_form.addRow("JPEG Quality", self.quality_spin)
        settings_card.body_layout.addLayout(settings_form)

        center_column = QtWidgets.QWidget()
        center_layout = QtWidgets.QVBoxLayout(center_column)
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.setSpacing(14)
        splitter.addWidget(center_column)

        rules_card = CardFrame(
            "Rules",
            "Use Add Rule for any available rule type. Starter rules can be edited, removed, or restored.",
        )
        center_layout.addWidget(rules_card)

        self.rule_list = QtWidgets.QListWidget()
        rules_card.body_layout.addWidget(self.rule_list)

        rule_buttons = QtWidgets.QHBoxLayout()
        self.add_rule_button = QtWidgets.QPushButton("Add Rule")
        self.remove_rule_button = QtWidgets.QPushButton("Remove Rule")
        self.reset_rules_button = QtWidgets.QPushButton("Restore Starter Rules")
        rule_buttons.addWidget(self.add_rule_button)
        rule_buttons.addWidget(self.remove_rule_button)
        rule_buttons.addWidget(self.reset_rules_button)
        rules_card.body_layout.addLayout(rule_buttons)

        right_column = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(right_column)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(14)
        splitter.addWidget(right_column)

        right_splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        right_splitter.setChildrenCollapsible(False)
        right_layout.addWidget(right_splitter, 1)

        editor_card = CardFrame(
            "Rule Settings",
            "Basic settings are shown first. Use the Advanced tab only when you need font and spacing control.",
        )
        editor_card.setMinimumHeight(360)
        right_splitter.addWidget(editor_card)

        self.rule_summary_label = QtWidgets.QLabel("Select a rule to change its settings.")
        self.rule_summary_label.setObjectName("CardSubtitle")
        self.rule_summary_label.setWordWrap(True)
        editor_card.body_layout.addWidget(self.rule_summary_label)

        self.editor_stack = QtWidgets.QStackedWidget()
        self.empty_editor = QtWidgets.QLabel("Select a rule from the list to edit it.")
        self.empty_editor.setAlignment(QtCore.Qt.AlignCenter)
        self.due_editor = DueDateRuleEditor()
        self.note_editor = NoteRuleEditor()
        self.rule_editors = {
            "due_date": self.due_editor,
            "note": self.note_editor,
        }
        self.editor_stack.addWidget(self.empty_editor)
        self.editor_stack.addWidget(self.due_editor)
        self.editor_stack.addWidget(self.note_editor)
        self.editor_stack.setMinimumHeight(360)
        editor_card.body_layout.addWidget(self.editor_stack)

        output_card = CardFrame("Step 3: Results", "Processed files appear here after each run.")
        output_card.setMinimumHeight(220)
        right_splitter.addWidget(output_card)

        self.output_list = QtWidgets.QListWidget()
        output_card.body_layout.addWidget(self.output_list)

        output_buttons = QtWidgets.QHBoxLayout()
        self.open_output_button = QtWidgets.QPushButton("Open Selected")
        self.open_output_folder_button = QtWidgets.QPushButton("Open Folder")
        output_buttons.addWidget(self.open_output_button)
        output_buttons.addWidget(self.open_output_folder_button)
        output_card.body_layout.addLayout(output_buttons)

        process_bar = QtWidgets.QHBoxLayout()
        self.process_button = QtWidgets.QPushButton("Process Files")
        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        process_bar.addWidget(self.process_button)
        process_bar.addWidget(self.progress_bar, 1)
        outer.addLayout(process_bar)

        self.log_output = QtWidgets.QPlainTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setMaximumBlockCount(400)
        self.log_output.setMinimumHeight(90)
        self.log_output.setMaximumHeight(120)
        self.log_output.setPlaceholderText("Processing details and errors will appear here.")
        outer.addWidget(self.log_output, 0)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        splitter.setStretchFactor(2, 2)
        splitter.setSizes([420, 420, 600])
        right_splitter.setSizes([520, 280])

        self.add_files_button.clicked.connect(self._add_files)
        self.remove_file_button.clicked.connect(self._remove_selected_files)
        self.clear_files_button.clicked.connect(self._clear_files)
        self.output_dir_button.clicked.connect(self._pick_output_dir)
        self.tesseract_button.clicked.connect(self._pick_tesseract)
        self.add_rule_button.clicked.connect(self._add_rule)
        self.remove_rule_button.clicked.connect(self._remove_selected_rule)
        self.reset_rules_button.clicked.connect(self._reset_rules)
        self.rule_list.currentRowChanged.connect(self._select_rule)
        self.process_button.clicked.connect(self._start_processing)
        self.open_output_button.clicked.connect(self._open_selected_output)
        self.open_output_folder_button.clicked.connect(self._open_output_folder)

        self.due_editor.changed.connect(self._sync_active_rule)
        self.note_editor.changed.connect(self._sync_active_rule)

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QWidget {
                background: #f3f5f8;
                color: #1f2d3a;
                font-family: "Segoe UI";
                font-size: 10.5pt;
            }
            QFrame#Card {
                background: #ffffff;
                border: 1px solid #d5dce5;
                border-radius: 12px;
            }
            QLabel#HeroTitle {
                font-size: 21pt;
                font-weight: 700;
                color: #17334d;
            }
            QLabel#HeroSubtitle {
                color: #60707f;
                font-size: 10pt;
                margin-bottom: 6px;
            }
            QLabel#CardTitle {
                font-size: 12.5pt;
                font-weight: 700;
                color: #17334d;
            }
            QLabel#CardSubtitle {
                color: #60707f;
                font-size: 9.5pt;
            }
            QPushButton {
                background: #1f5a92;
                color: white;
                border: none;
                border-radius: 9px;
                padding: 9px 14px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #2769a7;
            }
            QPushButton:disabled {
                background: #a8b3bf;
            }
            QListWidget, QLineEdit, QSpinBox, QComboBox, QPlainTextEdit, QFontComboBox {
                background: #ffffff;
                border: 1px solid #d1d9e2;
                border-radius: 9px;
                padding: 6px 8px;
                selection-background-color: #dbe8f5;
                selection-color: #17334d;
            }
            QProgressBar {
                border: 1px solid #d1d9e2;
                border-radius: 10px;
                background: #ffffff;
                text-align: center;
            }
            QProgressBar::chunk {
                background: #2c7a55;
                border-radius: 10px;
            }
            QTabWidget::pane {
                border: 1px solid #d5dce5;
                border-radius: 10px;
                background: #ffffff;
            }
            QTabBar::tab {
                background: #eef2f6;
                border: 1px solid #d5dce5;
                padding: 8px 14px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                margin-right: 4px;
            }
            QTabBar::tab:selected {
                background: #ffffff;
                color: #17334d;
            }
            """
        )

    def _load_state(self) -> None:
        if not STATE_PATH.exists():
            return
        try:
            payload = json.loads(STATE_PATH.read_text(encoding="utf-8"))
        except Exception:
            return

        self.input_files = payload.get("input_files", [])
        self.output_dir_edit.setText(payload.get("output_dir", self.output_dir_edit.text()))
        self.tesseract_edit.setText(payload.get("tesseract_path", self.tesseract_edit.text()))
        self.dpi_spin.setValue(payload.get("dpi", self.dpi_spin.value()))
        self.quality_spin.setValue(payload.get("jpeg_quality", self.quality_spin.value()))

        raw_rules = payload.get("rules", [])
        restored_rules = []
        for raw_rule in raw_rules:
            try:
                restored_rules.append(rule_from_dict(raw_rule))
            except Exception:
                continue
        if restored_rules:
            self.rules = restored_rules

    def _save_state(self) -> None:
        ensure_app_dir()
        payload = {
            "input_files": self.input_files,
            "output_dir": self.output_dir_edit.text().strip(),
            "tesseract_path": self.tesseract_edit.text().strip(),
            "dpi": self.dpi_spin.value(),
            "jpeg_quality": self.quality_spin.value(),
            "rules": [rule_to_dict(rule) for rule in self.rules],
        }
        STATE_PATH.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def _refresh_input_list(self) -> None:
        self.input_list.clear()
        for path in self.input_files:
            self.input_list.addItem(path)

    def _refresh_output_list(self) -> None:
        self.output_list.clear()
        for path in self.output_files:
            self.output_list.addItem(path)

    def _refresh_rule_list(self) -> None:
        current_row = self.rule_list.currentRow()
        self.rule_list.clear()
        for rule in self.rules:
            item = QtWidgets.QListWidgetItem(rule_list_text(rule))
            item.setToolTip(rule_summary(rule))
            self.rule_list.addItem(item)

        if self.rules:
            self.rule_list.setCurrentRow(max(0, min(current_row, len(self.rules) - 1)))
        else:
            self.rule_summary_label.setText("Select a rule to change its settings.")
            self.editor_stack.setCurrentWidget(self.empty_editor)

        self._update_rule_buttons()

    def _add_files(self) -> None:
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self,
            "Select Invoice PDFs",
            str(Path.cwd()),
            "PDF Files (*.pdf)",
        )
        if not files:
            return
        for path in files:
            if path not in self.input_files:
                self.input_files.append(path)
        self._refresh_input_list()
        self._save_state()

    def _remove_selected_files(self) -> None:
        rows = sorted({item.row() for item in self.input_list.selectedIndexes()}, reverse=True)
        for row in rows:
            self.input_files.pop(row)
        self._refresh_input_list()
        self._save_state()

    def _clear_files(self) -> None:
        self.input_files = []
        self._refresh_input_list()
        self._save_state()

    def _pick_output_dir(self) -> None:
        directory = QtWidgets.QFileDialog.getExistingDirectory(
            self,
            "Choose Output Folder",
            self.output_dir_edit.text().strip() or str(Path.cwd()),
        )
        if directory:
            self.output_dir_edit.setText(directory)
            self._save_state()

    def _pick_tesseract(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Locate Tesseract",
            self.tesseract_edit.text().strip() or str(Path(r"C:\Program Files")),
            "Executable (*.exe)",
        )
        if path:
            self.tesseract_edit.setText(path)
            self._save_state()

    def _count_rules(self, definition: RuleDefinition) -> int:
        return sum(isinstance(rule, definition.config_type) for rule in self.rules)

    def _available_rule_definitions(self) -> List[RuleDefinition]:
        available = []
        for definition in RULE_DEFINITIONS:
            if definition.max_instances is None or self._count_rules(definition) < definition.max_instances:
                available.append(definition)
        return available

    def _update_rule_buttons(self) -> None:
        available = self._available_rule_definitions()
        self.add_rule_button.setEnabled(bool(available))
        self.remove_rule_button.setEnabled(self.rule_list.currentRow() >= 0)

    def _add_rule(self) -> None:
        available = self._available_rule_definitions()
        if not available:
            QtWidgets.QMessageBox.information(
                self,
                APP_TITLE,
                "All available rule types are already in the list.",
            )
            return

        dialog = AddRuleDialog(available, self)
        if dialog.exec() != QtWidgets.QDialog.Accepted:
            return

        definition = dialog.selected_definition()
        if definition is None:
            return

        self.rules.append(definition.factory())
        self._refresh_rule_list()
        self.rule_list.setCurrentRow(len(self.rules) - 1)
        self._save_state()

    def _remove_selected_rule(self) -> None:
        row = self.rule_list.currentRow()
        if row < 0:
            return
        self.rules.pop(row)
        self._refresh_rule_list()
        self._save_state()

    def _reset_rules(self) -> None:
        self.rules = starter_rules()
        self._refresh_rule_list()
        self.rule_list.setCurrentRow(0)
        self._save_state()

    def _select_rule(self, row: int) -> None:
        if row < 0 or row >= len(self.rules):
            self.rule_summary_label.setText("Select a rule to change its settings.")
            self.editor_stack.setCurrentWidget(self.empty_editor)
            return
        rule = self.rules[row]
        definition = rule_definition_for_rule(rule)
        editor = self.rule_editors.get(definition.editor_key)
        if editor is None:
            self.rule_summary_label.setText("Select a rule to change its settings.")
            self.editor_stack.setCurrentWidget(self.empty_editor)
            return

        if isinstance(editor, DueDateRuleEditor):
            editor.set_rule(rule)
        elif isinstance(editor, NoteRuleEditor):
            editor.set_rule(rule)
        self.rule_summary_label.setText(rule_summary(rule))
        self.editor_stack.setCurrentWidget(editor)
        self._update_rule_buttons()

    def _sync_active_rule(self) -> None:
        row = self.rule_list.currentRow()
        if row < 0 or row >= len(self.rules):
            return

        rule = self.rules[row]
        definition = rule_definition_for_rule(rule)
        editor = self.rule_editors.get(definition.editor_key)
        if isinstance(editor, DueDateRuleEditor):
            editor.apply_changes()
        elif isinstance(editor, NoteRuleEditor):
            editor.apply_changes()

        current_item = self.rule_list.item(row)
        if current_item is not None:
            current_item.setText(rule_list_text(rule))
            current_item.setToolTip(rule_summary(rule))
        self.rule_summary_label.setText(rule_summary(rule))
        self._save_state()

    def _append_log(self, message: str) -> None:
        self.log_output.appendPlainText(message)
        append_app_log(message)

    def _validate_processing_request(self) -> Optional[str]:
        if not self.input_files:
            return "Add at least one PDF before processing."

        missing_inputs = [path for path in self.input_files if not Path(path).exists()]
        if missing_inputs:
            preview = "\n".join(missing_inputs[:5])
            if len(missing_inputs) > 5:
                preview += "\n..."
            return "Missing input file(s):\n%s" % preview

        if not any(getattr(rule, "enabled", True) for rule in self.rules):
            return "Turn on at least one rule before processing."

        output_dir = self.output_dir_edit.text().strip()
        if not output_dir:
            return "Choose an output folder before processing."

        try:
            backend.configure_tesseract(self.tesseract_edit.text().strip() or None)
        except Exception as exc:
            return str(exc)

        return None

    def _start_processing(self) -> None:
        validation_error = self._validate_processing_request()
        if validation_error:
            self._append_log(validation_error)
            QtWidgets.QMessageBox.warning(self, APP_TITLE, validation_error)
            return

        self.process_button.setEnabled(False)
        self.progress_bar.setValue(0)
        self.log_output.clear()
        self._append_log("Starting processing for %d file(s)." % len(self.input_files))
        self._save_state()

        self.processing_thread = QtCore.QThread(self)
        self.worker = ProcessWorker(
            input_files=self.input_files,
            output_dir=self.output_dir_edit.text().strip(),
            dpi=self.dpi_spin.value(),
            jpeg_quality=self.quality_spin.value(),
            tesseract_path=self.tesseract_edit.text().strip(),
            rules=clone_rules(self.rules),
        )
        self.worker.moveToThread(self.processing_thread)
        self.processing_thread.started.connect(self.worker.run)
        self.worker.log.connect(self._append_log)
        self.worker.progress.connect(self._update_progress)
        self.worker.finished.connect(self._processing_finished)
        self.worker.failed.connect(self._processing_failed)
        self.worker.finished.connect(self.processing_thread.quit)
        self.worker.failed.connect(self.processing_thread.quit)
        self.processing_thread.finished.connect(self._cleanup_worker)
        self.processing_thread.start()

    def _update_progress(self, current: int, total: int) -> None:
        value = 0 if total == 0 else int((current / total) * 100)
        self.progress_bar.setValue(value)

    def _processing_finished(self, outputs: List[str]) -> None:
        self.process_button.setEnabled(True)
        self.progress_bar.setValue(100)
        self.output_files = outputs
        self._refresh_output_list()
        self._append_log("All files processed.")

    def _processing_failed(self, trace_text: str) -> None:
        self.process_button.setEnabled(True)
        self.progress_bar.setValue(0)
        self._append_log(trace_text)
        QtWidgets.QMessageBox.critical(self, APP_TITLE, "Processing failed. See the log for details.")

    def _cleanup_worker(self) -> None:
        if self.worker is not None:
            self.worker.deleteLater()
            self.worker = None
        if self.processing_thread is not None:
            self.processing_thread.deleteLater()
            self.processing_thread = None

    def _open_selected_output(self) -> None:
        item = self.output_list.currentItem()
        if not item:
            return
        path = item.text()
        if hasattr(os, "startfile"):
            os.startfile(path)

    def _open_output_folder(self) -> None:
        folder = self.output_dir_edit.text().strip()
        if folder and hasattr(os, "startfile"):
            os.startfile(folder)

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        self._save_state()
        append_app_log("Application closed.")
        super().closeEvent(event)


def main() -> int:
    ensure_app_dir()
    append_app_log("Application launched.")
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationDisplayName(APP_TITLE)
    app.setOrganizationName("JLJ IV Enterprises Inc.")
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
