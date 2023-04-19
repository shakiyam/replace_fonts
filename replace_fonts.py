import argparse
import os.path
import shutil
from datetime import datetime
from typing import Optional

from lxml.etree import _Element

from pptx import Presentation
from pptx.oxml import CT_TextCharacterProperties
from pptx.oxml.ns import qn
from pptx.oxml.text import CT_TextFont
from pptx.shapes.base import BaseShape
from pptx.shapes.group import GroupShape
from pptx.text.text import TextFrame

version = '2023-04-19'


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


def replace_latin_font(latin: CT_TextFont, major_or_minor: str, text: Optional[str] = None) -> None:
    if major_or_minor == 'major':
        default_font = '+mj-lt'
    else:
        default_font = '+mn-lt'
    latin_font = latin.get('typeface')
    if args.code and latin_font == 'Consolas':
        log(f'Preserve {major_or_minor} latin font as {latin_font}', text)
    elif args.code and latin_font == 'Courier New':
        latin.set('typeface', 'Consolas')
        log(f'Replace {major_or_minor} latin font from {latin_font} to Consolas', text)
    elif latin_font != default_font:
        latin.set('typeface', default_font)
        log(f'Replace {major_or_minor} latin font from {latin_font} to {default_font}', text)


def replace_ea_font(ea: _Element, major_or_minor: str, text: Optional[str] = None) -> None:
    if major_or_minor == 'major':
        default_font = '+mj-ea'
    else:
        default_font = '+mn-ea'
    ea_font = ea.get('typeface')
    if args.code and ea_font == 'Consolas':
        log(f'Preserve {major_or_minor} east asian font as {ea_font}', text)
    elif args.code and ea_font == 'Courier New':
        ea.set('typeface', 'Consolas')
        log(f'Replace {major_or_minor} east asian font from {ea_font} to Consolas', text)
    elif ea_font != default_font:
        ea.set('typeface', default_font)
        log(f'Replace {major_or_minor} east asian font from {ea_font} to {default_font}', text)


def replace_properties_fonts(pr: CT_TextCharacterProperties, major_or_minor: str, text: Optional[str] = None) -> None:
    if pr.find(qn('a:latin')) is not None:
        replace_latin_font(pr.find(qn('a:latin')), major_or_minor, text)
    if pr.find(qn('a:ea')) is not None:
        replace_ea_font(pr.find(qn('a:ea')), major_or_minor, text)


def replace_text_frame_fonts(text_frame: TextFrame, major_or_minor: str) -> None:
    for paragraph in text_frame.paragraphs:
        if paragraph._element.pPr is not None and paragraph._element.pPr.defRPr is not None:
            replace_properties_fonts(paragraph._element.pPr.defRPr, major_or_minor)
        for run in paragraph.runs:
            text = run.text.strip()
            replace_properties_fonts(run.font._element, major_or_minor, text)
        if paragraph._element.endParaRPr is not None:
            replace_properties_fonts(paragraph._element.endParaRPr, major_or_minor)


def replace_shape_fonts(shape: BaseShape) -> None:
    if shape.has_text_frame:
        ph = shape.element.find(f".//{qn('p:ph')}")
        if ph is not None and ph.get('type') in ['ctrTitle', 'title']:
            replace_text_frame_fonts(shape.text_frame, 'major')
        else:
            replace_text_frame_fonts(shape.text_frame, 'minor')
    elif shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                replace_text_frame_fonts(cell.text_frame, 'minor')
    elif isinstance(shape, GroupShape):
        for item in shape.shapes:
            replace_shape_fonts(item)
    elif shape.has_chart:
        for latin in shape.chart.element.findall(f".//{qn('a:latin')}"):
            replace_latin_font(latin, 'minor')
        for ea in shape.chart.element.findall(f".//{qn('a:ea')}"):
            replace_ea_font(ea, 'minor')


print(f'replace_fonts - version {version} by Shinichi Akiyama')

parser = argparse.ArgumentParser()
parser.add_argument('files', nargs='*')
parser.add_argument('--code', help='keep fonts of the code', action='store_true')
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
                print(tx_style.tag)
                if tx_style.tag == qn('p:titleStyle'):
                    major_or_minor = 'major'
                else:
                    major_or_minor = 'minor'
                for list_style in tx_style.getchildren():
                    if isinstance(list_style, CT_TextCharacterProperties):
                        replace_properties_fonts(list_style, major_or_minor)
                    else:
                        replace_properties_fonts(list_style.find(qn('a:defRPr')), major_or_minor, list_style.tag)

        presentation.save(file)
        log(f'{file} was saved.')

print('All files were processed.')
