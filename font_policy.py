from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import TextIO

import yaml
from lxml import etree
from lxml.etree import _Element
from pptx.presentation import Presentation as PresentationType

EA_SCRIPTS = ("Jpan", "Hang", "Hans", "Hant")

THEME_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.theme+xml"


@dataclass(frozen=True)
class FontPolicy:
    major_latin: str
    major_ea: str
    minor_latin: str
    minor_ea: str


def load_font_policy(path: Path) -> FontPolicy:
    with open(path) as f:
        data = yaml.safe_load(f)
    if not isinstance(data, dict):
        msg = "Font policy must be a YAML mapping"
        raise ValueError(msg)
    missing = []
    for level in ("major", "minor"):
        if not isinstance(data.get("theme_fonts", {}).get(level), dict):
            missing.extend([f"theme_fonts.{level}.latin", f"theme_fonts.{level}.ea"])
            continue
        for key in ("latin", "ea"):
            if key not in data["theme_fonts"][level]:
                missing.append(f"theme_fonts.{level}.{key}")
    if missing:
        msg = f"Font policy missing required keys: {', '.join(missing)}"
        raise ValueError(msg)
    tf = data["theme_fonts"]
    return FontPolicy(
        major_latin=tf["major"]["latin"],
        major_ea=tf["major"]["ea"],
        minor_latin=tf["minor"]["latin"],
        minor_ea=tf["minor"]["ea"],
    )


LogFn = Callable[[TextIO, str], None]


def _update_theme_element(
    element: _Element | None,
    new_val: str,
    label: str,
    log_file: TextIO,
    log_fn: LogFn,
) -> None:
    if element is None:
        return
    old = element.get("typeface")
    if old != new_val:
        log_fn(
            log_file,
            f'Update theme {label} from "{old}" to "{new_val}"',
        )
        element.set("typeface", new_val)


def update_theme_fonts(
    presentation: PresentationType,
    policy: FontPolicy,
    log_file: TextIO,
    log_fn: LogFn,
) -> None:
    a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    font_configs = [
        ("major", f"{{{a_ns}}}majorFont", policy.major_latin, policy.major_ea),
        ("minor", f"{{{a_ns}}}minorFont", policy.minor_latin, policy.minor_ea),
    ]
    for part in presentation.part.package.iter_parts():
        if part.content_type != THEME_CONTENT_TYPE:
            continue
        root = etree.fromstring(part.blob)
        for level_name, font_tag, latin_val, ea_val in font_configs:
            font_group = root.find(f".//{font_tag}")
            if font_group is None:
                continue
            _update_theme_element(
                font_group.find(f"{{{a_ns}}}latin"),
                latin_val, f"{level_name} latin", log_file, log_fn,
            )
            _update_theme_element(
                font_group.find(f"{{{a_ns}}}ea"),
                ea_val, f"{level_name} ea", log_file, log_fn,
            )
            for script in EA_SCRIPTS:
                el = font_group.find(
                    f"{{{a_ns}}}font[@script='{script}']",
                )
                _update_theme_element(
                    el, ea_val,
                    f"{level_name} ea script {script}",
                    log_file, log_fn,
                )
        part._blob = etree.tostring(
            root, xml_declaration=True,
            encoding="UTF-8", standalone=True,
        )
