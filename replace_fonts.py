import argparse
import os.path
import shutil
from datetime import datetime
from enum import Enum
from typing import Optional

from lxml.etree import _Element

from pptx import Presentation
from pptx.oxml import CT_TextCharacterProperties
from pptx.oxml.ns import qn
from pptx.shapes.autoshape import Shape
from pptx.shapes.base import BaseShape
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.group import GroupShape
from pptx.text.text import TextFrame

__version__ = '2025-10-17'


class FontType(Enum):
    MAJOR = 'major'
    MINOR = 'minor'


class FontCategory(Enum):
    LATIN = 'latin font'
    EAST_ASIAN = 'east asian font'


FONT_MAPPINGS = {
    FontCategory.LATIN: {
        FontType.MAJOR: '+mj-lt',
        FontType.MINOR: '+mn-lt',
    },
    FontCategory.EAST_ASIAN: {
        FontType.MAJOR: '+mj-ea',
        FontType.MINOR: '+mn-ea',
    },
}

PRESERVED_CODE_FONT = 'Consolas'
REPLACED_CODE_FONTS = ('Courier New',)


def log(message: str, text: Optional[str] = None) -> None:
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if text is not None:
        message = f'[{text}] {message}'
    print(f'{timestamp} {message}', file=logfile)
    print(f'{timestamp} {message}')


def backup_file(path: str) -> str:
    base, ext = os.path.splitext(path)
    backup = f'{base} - backup{ext}'
    num = 2
    while os.path.exists(backup):
        backup = f'{base} - backup ({num}){ext}'
        num += 1
    shutil.copyfile(path, backup)
    return backup


def replace_font_element(element: _Element, font_type: FontType, font_category: FontCategory,
                         text: Optional[str] = None) -> None:
    default_font = FONT_MAPPINGS[font_category][font_type]
    current_font = element.get('typeface')
    if args.code and current_font == PRESERVED_CODE_FONT:
        log(f'Preserve {font_type.value} {font_category.value} as {current_font}', text)
    elif args.code and current_font in REPLACED_CODE_FONTS:
        element.set('typeface', PRESERVED_CODE_FONT)
        log(f'Replace {font_type.value} {font_category.value} from {current_font} to {PRESERVED_CODE_FONT}', text)
    elif current_font != default_font:
        element.set('typeface', default_font)
        log(f'Replace {font_type.value} {font_category.value} from {current_font} to {default_font}', text)


def replace_properties_fonts(properties: CT_TextCharacterProperties, font_type: FontType, text: Optional[str] = None) -> None:
    if properties.find(qn('a:latin')) is not None:
        replace_font_element(properties.find(qn('a:latin')), font_type, FontCategory.LATIN, text)
    if properties.find(qn('a:ea')) is not None:
        replace_font_element(properties.find(qn('a:ea')), font_type, FontCategory.EAST_ASIAN, text)


def replace_text_frame_fonts(text_frame: TextFrame, font_type: FontType) -> None:
    for paragraph in text_frame.paragraphs:
        if paragraph._element.pPr is not None and paragraph._element.pPr.defRPr is not None:
            replace_properties_fonts(paragraph._element.pPr.defRPr, font_type)
        for run in paragraph.runs:
            text = run.text.strip()
            replace_properties_fonts(run.font._element, font_type, text)
        if paragraph._element.endParaRPr is not None:
            replace_properties_fonts(paragraph._element.endParaRPr, font_type)


def replace_shape_fonts(shape: BaseShape) -> None:
    if isinstance(shape, Shape):
        ph = shape.element.find(f".//{qn('p:ph')}")
        if ph is not None and ph.get('type') in ['ctrTitle', 'title']:
            replace_text_frame_fonts(shape.text_frame, FontType.MAJOR)
        else:
            replace_text_frame_fonts(shape.text_frame, FontType.MINOR)
    elif isinstance(shape, GraphicFrame) and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                replace_text_frame_fonts(cell.text_frame, FontType.MINOR)
    elif isinstance(shape, GraphicFrame) and shape.has_chart:
        for latin in shape.chart.element.findall(f".//{qn('a:latin')}"):
            replace_font_element(latin, FontType.MINOR, FontCategory.LATIN)
        for east_asian in shape.chart.element.findall(f".//{qn('a:ea')}"):
            replace_font_element(east_asian, FontType.MINOR, FontCategory.EAST_ASIAN)
    elif isinstance(shape, GroupShape):
        for item in shape.shapes:
            replace_shape_fonts(item)


print(f'replace_fonts - version {__version__} by Shinichi Akiyama')

parser = argparse.ArgumentParser()
parser.add_argument('files', nargs='*')
parser.add_argument('--code', help='preserve code fonts', action='store_true')
args = parser.parse_args()

for file in args.files:
    base, ext = os.path.splitext(file)
    with open(f'{base}.log', 'a') as logfile:
        backup = backup_file(file)
        log(f'{file} was backed up to {backup}.')

        presentation = Presentation(file)
        log(f'{file} was opened.')

        for i, slide in enumerate(presentation.slides):
            log(f'--- Slide {i + 1} ---')
            for shape in slide.shapes:
                replace_shape_fonts(shape)
        for i, slide_master in enumerate(presentation.slide_masters):
            log(f'--- Slide Master {i + 1} ---')
            tx_styles = slide_master.element.find(qn('p:txStyles'))
            for tx_style in tx_styles.getchildren():
                if tx_style.tag == qn('p:titleStyle'):
                    font_type = FontType.MAJOR
                else:
                    font_type = FontType.MINOR
                for list_style in tx_style.getchildren():
                    if isinstance(list_style, CT_TextCharacterProperties):
                        replace_properties_fonts(list_style, font_type)
                    else:
                        replace_properties_fonts(list_style.find(qn('a:defRPr')), font_type, list_style.tag)

        presentation.save(file)
        log(f'{file} was saved.')

print('All files were processed.')
