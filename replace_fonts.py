import argparse
import os.path
import shutil
from datetime import datetime
from typing import Optional

from lxml import etree

from pptx import Presentation
from pptx.oxml import CT_TextCharacterProperties
from pptx.oxml.ns import qn
from pptx.shapes.base import BaseShape
from pptx.shapes.group import GroupShape
from pptx.text.text import TextFrame

version = '2023-03-27'


def log(message: str) -> None:
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
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


def replace_properties_fonts(pr: CT_TextCharacterProperties, major_or_minor: str, text: Optional[str]) -> None:
    if major_or_minor == 'major':
        default_font = {'latin': '+mj-lt', 'east asian': '+mj-ea'}
    else:
        default_font = {'latin': '+mn-lt', 'east asian': '+mn-ea'}
    if pr.find(qn('a:latin')) is not None:
        latin_font = pr.find(qn('a:latin')).get('typeface')
        if args.code and latin_font == 'Consolas':
            if text is not None:
                log(f'[{text}] Keep {major_or_minor} latin font as {latin_font}')
            else:
                log(f'Keep {major_or_minor} latin font as {latin_font}')
        elif args.code and latin_font == 'Courier New':
            pr.find(qn('a:latin')).set('typeface', 'Consolas')
            if text is not None:
                log(f'[{text}] Replace {major_or_minor} latin font from {latin_font} to Consolas')
            else:
                log(f'Replace {major_or_minor} latin font from {latin_font} to Consolas')
        else:
            etree.strip_elements(pr, qn('a:latin'))
            if text is not None:
                log(f"[{text}] Replace {major_or_minor} latin font from {latin_font} to {default_font['latin']}")
            else:
                log(f"Replace {major_or_minor} latin font from {latin_font} to  {default_font['latin']}")
    if pr.find(qn('a:ea')) is not None:
        ea_font = pr.find(qn('a:ea')).get('typeface')
        etree.strip_elements(pr, qn('a:ea'))
        if text is not None:
            log(f"[{text}] Replace {major_or_minor} east asian font from {ea_font} to {default_font['east asian']}")
        else:
            log(f"Replace {major_or_minor} east asian font from {ea_font} to {default_font['east asian']}")


def replace_text_frame_fonts(text_frame: TextFrame, major_or_minor: str) -> None:
    for paragraph in text_frame.paragraphs:
        if paragraph._element.pPr is not None and paragraph._element.pPr.defRPr is not None:
            replace_properties_fonts(paragraph._element.pPr.defRPr, major_or_minor, None)
        for run in paragraph.runs:
            text = run.text.strip()
            replace_properties_fonts(run.font._element, major_or_minor, text)
        if paragraph._element.endParaRPr is not None:
            replace_properties_fonts(paragraph._element.endParaRPr, major_or_minor, None)


def replace_shape_fonts(shape: BaseShape) -> None:
    if shape.has_text_frame:
        ph = shape.element.find('.//{*}ph')
        if ph is not None and ph.get('type') == 'title':
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
        presentation.save(file)
        log(f'{file} was saved.')

print('All files were processed.')
