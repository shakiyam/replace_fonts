import argparse
import os.path
import shutil
from datetime import datetime
from typing import Optional

from pptx import Presentation
from pptx.oxml import CT_TextCharacterProperties
from pptx.oxml.ns import qn
from pptx.shapes.base import BaseShape
from pptx.shapes.group import GroupShape
from pptx.text.text import TextFrame

version = '2023-03-28'


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


def replace_properties_fonts(pr: CT_TextCharacterProperties, major_or_minor: str, text: Optional[str] = None) -> None:
    if major_or_minor == 'major':
        default_font = {'latin': '+mj-lt', 'east asian': '+mj-ea'}
    else:
        default_font = {'latin': '+mn-lt', 'east asian': '+mn-ea'}
    if pr.find(qn('a:latin')) is not None:
        latin_font = pr.find(qn('a:latin')).get('typeface')
        if args.code and latin_font == 'Consolas':
            log(f'Keep {major_or_minor} latin font as {latin_font}', text)
        elif args.code and latin_font == 'Courier New':
            pr.find(qn('a:latin')).set('typeface', 'Consolas')
            log(f'Replace {major_or_minor} latin font from {latin_font} to Consolas', text)
        elif latin_font != default_font['latin']:
            pr.find(qn('a:latin')).set('typeface', default_font['latin'])
            log(f"Replace {major_or_minor} latin font from {latin_font} to {default_font['latin']}", text)
    if pr.find(qn('a:ea')) is not None:
        ea_font = pr.find(qn('a:ea')).get('typeface')
        if args.code and ea_font == 'Consolas':
            log(f'Keep {major_or_minor} east asian font as {ea_font}', text)
        elif args.code and ea_font == 'Courier New':
            pr.find(qn('a:ea')).set('typeface', 'Consolas')
            log(f'Replace {major_or_minor} east asian font from {ea_font} to Consolas', text)
        elif ea_font != default_font['east asian']:
            pr.find(qn('a:ea')).set('typeface', default_font['east asian'])
            log(f"Replace {major_or_minor} east asian font from {ea_font} to {default_font['east asian']}", text)


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
