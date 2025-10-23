import argparse
import os.path
import shutil
from datetime import datetime
from enum import Enum
from typing import Optional, TextIO

from lxml.etree import _Element

from pptx import Presentation
from pptx.exc import PackageNotFoundError
from pptx.oxml import CT_TextCharacterProperties  # type: ignore[attr-defined]
from pptx.oxml.ns import qn
from pptx.presentation import Presentation as PresentationType
from pptx.shapes.autoshape import Shape
from pptx.shapes.base import BaseShape
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.group import GroupShape
from pptx.slide import SlideMasters, Slides
from pptx.text.text import TextFrame

__version__ = '2025-10-22'


class ThemeFont(Enum):
    MAJOR = 'major'
    MINOR = 'minor'


class FontScript(Enum):
    LATIN = 'latin'
    EAST_ASIAN = 'east asian'


FONT_MAPPINGS = {
    FontScript.LATIN: {
        ThemeFont.MAJOR: '+mj-lt',
        ThemeFont.MINOR: '+mn-lt',
    },
    FontScript.EAST_ASIAN: {
        ThemeFont.MAJOR: '+mj-ea',
        ThemeFont.MINOR: '+mn-ea',
    },
}

FONT_ELEMENT_MAPPINGS = [
    (qn('a:latin'), FontScript.LATIN),
    (qn('a:ea'), FontScript.EAST_ASIAN),
]

PRESERVED_CODE_FONT = 'Consolas'
REPLACED_CODE_FONTS = ('Courier New',)


def log(log_file: TextIO, message: str, element_text: Optional[str] = None) -> None:
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if element_text is not None:
        message = f'[{element_text}] {message}'
    print(f'{timestamp} {message}', file=log_file)
    print(f'{timestamp} {message}')


def log_font_action(
    theme_font: ThemeFont,
    font_script: FontScript,
    current_font: str,
    new_font: Optional[str],
    log_file: TextIO,
    element_text: Optional[str] = None,
) -> None:
    if new_font:
        message = (
            f'Replace {theme_font.value} {font_script.value} '
            f'from {current_font} to {new_font}'
        )
    else:
        message = f'Preserve {theme_font.value} {font_script.value} as {current_font}'
    log(log_file, message, element_text)


def create_backup(path: str) -> str:
    base, ext = os.path.splitext(path)
    backup_path = f'{base} - backup{ext}'
    backup_number = 2
    while os.path.exists(backup_path):
        backup_path = f'{base} - backup ({backup_number}){ext}'
        backup_number += 1
    shutil.copyfile(path, backup_path)
    return backup_path


def replace_font_element(
    element: _Element,
    theme_font: ThemeFont,
    font_script: FontScript,
    preserve_code_fonts: bool,
    log_file: TextIO,
    element_text: Optional[str] = None,
) -> None:
    default_font = FONT_MAPPINGS[font_script][theme_font]
    current_font = element.get('typeface')
    if preserve_code_fonts and current_font == PRESERVED_CODE_FONT:
        log_font_action(
            theme_font, font_script, current_font, None, log_file, element_text
        )
    elif preserve_code_fonts and current_font in REPLACED_CODE_FONTS:
        element.set('typeface', PRESERVED_CODE_FONT)
        log_font_action(
            theme_font,
            font_script,
            current_font,
            PRESERVED_CODE_FONT,
            log_file,
            element_text,
        )
    elif current_font != default_font:
        element.set('typeface', default_font)
        log_font_action(
            theme_font, font_script, current_font, default_font, log_file, element_text
        )


def replace_properties_fonts(
    properties: CT_TextCharacterProperties,
    theme_font: ThemeFont,
    preserve_code_fonts: bool,
    log_file: TextIO,
    element_text: Optional[str] = None,
) -> None:
    for qname, font_script in FONT_ELEMENT_MAPPINGS:
        element = properties.find(qname)
        if element is not None:
            replace_font_element(
                element,
                theme_font,
                font_script,
                preserve_code_fonts,
                log_file,
                element_text,
            )


def replace_text_frame_fonts(
    text_frame: TextFrame,
    theme_font: ThemeFont,
    preserve_code_fonts: bool,
    log_file: TextIO,
) -> None:
    for paragraph in text_frame.paragraphs:
        if (
            paragraph._element.pPr is not None
            and paragraph._element.pPr.defRPr is not None
        ):
            replace_properties_fonts(
                paragraph._element.pPr.defRPr,
                theme_font,
                preserve_code_fonts,
                log_file,
            )
        for run in paragraph.runs:
            run_text = run.text.strip()
            replace_properties_fonts(
                run.font._element,
                theme_font,
                preserve_code_fonts,
                log_file,
                run_text,
            )
        if paragraph._element.endParaRPr is not None:
            replace_properties_fonts(
                paragraph._element.endParaRPr,
                theme_font,
                preserve_code_fonts,
                log_file,
            )


def replace_shape_text_fonts(
    shape: Shape, preserve_code_fonts: bool, log_file: TextIO
) -> None:
    placeholder = shape.element.find(f".//{qn('p:ph')}")
    if placeholder is not None and placeholder.get('type') in ['ctrTitle', 'title']:
        theme_font = ThemeFont.MAJOR
    else:
        theme_font = ThemeFont.MINOR
    replace_text_frame_fonts(
        shape.text_frame, theme_font, preserve_code_fonts, log_file
    )


def replace_table_fonts(
    shape: GraphicFrame, preserve_code_fonts: bool, log_file: TextIO
) -> None:
    for row in shape.table.rows:
        for cell in row.cells:
            replace_text_frame_fonts(
                cell.text_frame, ThemeFont.MINOR, preserve_code_fonts, log_file
            )


def replace_chart_fonts(
    shape: GraphicFrame, preserve_code_fonts: bool, log_file: TextIO
) -> None:
    for qname, font_script in FONT_ELEMENT_MAPPINGS:
        for element in shape.chart.element.findall(f'.//{qname}'):
            replace_font_element(
                element, ThemeFont.MINOR, font_script, preserve_code_fonts, log_file
            )


def replace_group_fonts(
    shape: GroupShape, preserve_code_fonts: bool, log_file: TextIO
) -> None:
    for item in shape.shapes:
        replace_shape_fonts(item, preserve_code_fonts, log_file)


def replace_shape_fonts(
    shape: BaseShape, preserve_code_fonts: bool, log_file: TextIO
) -> None:
    if isinstance(shape, Shape):
        replace_shape_text_fonts(shape, preserve_code_fonts, log_file)
    elif isinstance(shape, GraphicFrame) and shape.has_table:
        replace_table_fonts(shape, preserve_code_fonts, log_file)
    elif isinstance(shape, GraphicFrame) and shape.has_chart:
        replace_chart_fonts(shape, preserve_code_fonts, log_file)
    elif isinstance(shape, GroupShape):
        replace_group_fonts(shape, preserve_code_fonts, log_file)


def process_slides(
    slides: Slides, preserve_code_fonts: bool, log_file: TextIO
) -> None:
    for i, slide in enumerate(slides):
        log(log_file, f'--- Slide {i + 1} ---')
        for shape in slide.shapes:
            replace_shape_fonts(shape, preserve_code_fonts, log_file)


def process_slide_masters(
    slide_masters: SlideMasters, preserve_code_fonts: bool, log_file: TextIO
) -> None:
    for i, slide_master in enumerate(slide_masters):
        log(log_file, f'--- Slide Master {i + 1} ---')
        text_styles = slide_master.element.find(qn('p:txStyles'))
        if text_styles is None:
            continue
        for text_style in text_styles:
            if text_style.tag == qn('p:titleStyle'):
                theme_font = ThemeFont.MAJOR
            else:
                theme_font = ThemeFont.MINOR
            for list_style in text_style:
                if isinstance(list_style, CT_TextCharacterProperties):
                    replace_properties_fonts(
                        list_style,
                        theme_font,
                        preserve_code_fonts,
                        log_file,
                    )
                else:
                    def_rpr = list_style.find(qn('a:defRPr'))
                    if def_rpr is None:
                        continue
                    replace_properties_fonts(
                        def_rpr,
                        theme_font,
                        preserve_code_fonts,
                        log_file,
                    )


def process_presentation(
    presentation: PresentationType, preserve_code_fonts: bool, log_file: TextIO
) -> None:
    process_slides(presentation.slides, preserve_code_fonts, log_file)
    process_slide_masters(presentation.slide_masters, preserve_code_fonts, log_file)


def process_pptx_file(pptx_path: str, preserve_code_fonts: bool) -> None:
    base, _ = os.path.splitext(pptx_path)
    log_path = f'{base}.log'
    with open(log_path, 'a') as log_file:
        backup_path = create_backup(pptx_path)
        log(log_file, f'{pptx_path} was backed up to {backup_path}.')

        presentation = Presentation(pptx_path)
        log(log_file, f'{pptx_path} was opened.')

        process_presentation(presentation, preserve_code_fonts, log_file)

        presentation.save(pptx_path)
        log(log_file, f'{pptx_path} was saved.')


def main() -> int:
    print(f'replace_fonts - version {__version__} by Shinichi Akiyama')

    parser = argparse.ArgumentParser(
        description='Replace fonts in PowerPoint presentations'
    )
    parser.add_argument(
        'files', nargs='*', metavar='FILE', help='PowerPoint (.pptx) files to process'
    )
    parser.add_argument('--code', help='preserve code fonts', action='store_true')
    args = parser.parse_args()
    preserve_code_fonts = args.code

    if not args.files:
        print('No files specified.')
        return 0

    success_count = 0
    failure_count = 0

    for pptx_path in args.files:
        try:
            process_pptx_file(pptx_path, preserve_code_fonts)
            success_count += 1
        except (FileNotFoundError, PackageNotFoundError):
            print(f'Error: File not found or invalid: {pptx_path}')
            failure_count += 1
        except Exception as e:
            print(f'Error processing {pptx_path}: {type(e).__name__}: {e}')
            failure_count += 1

    total = success_count + failure_count
    if failure_count > 0:
        print(
            f'Processing complete: {success_count} succeeded, '
            f'{failure_count} failed out of {total}.'
        )
    else:
        print(f'All {total} file(s) processed successfully.')

    return 1 if failure_count > 0 else 0


if __name__ == '__main__':
    exit(main())
