import os.path
import shutil
import sys
from datetime import datetime
from typing import Optional

from lxml import etree

from pptx import Presentation
from pptx.oxml import CT_TextCharacterProperties
from pptx.oxml.ns import qn
from pptx.shapes.base import BaseShape
from pptx.shapes.group import GroupShape
from pptx.text.text import TextFrame

version = '2023-03-22'


def log(message: str):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f'{timestamp} {message}', file=logfile)
    print(f'{timestamp} {message}')


def backup_file(path):
    base, ext = os.path.splitext(path)
    backup = f'{base} - backup{ext}'
    num = 2
    while os.path.exists(backup):
        backup = f'{base} - backup ({num}){ext}'
        num += 1
    shutil.copyfile(path, backup)
    return backup


def replace_properties_fonts(pr: CT_TextCharacterProperties, is_major: bool, text: Optional[str]):
    if pr.find(qn('a:latin')) is not None:
        latin_font_name = pr.find(qn('a:latin')).get('typeface')
        etree.strip_elements(pr, qn('a:latin'))
        # run.font._rpr.latin.typeface = 'Meiryo UI'
        if is_major:
            if text is not None:
                log(f'[{text}] Major latin font has been converted from {latin_font_name} to +mj-lt')
            else:
                log(f'Major latin font has been converted from {latin_font_name} to +mj-lt')
        else:
            if text is not None:
                log(f'[{text}] Minor latin font has been converted from {latin_font_name} to +mn-lt')
            else:
                log(f'Minor latin font has been converted from {latin_font_name} to +mn-lt')
    if pr.find(qn('a:ea')) is not None:
        ea_font_name = pr.find(qn('a:ea')).get('typeface')
        etree.strip_elements(pr, qn('a:ea'))
        # ea = etree.SubElement(rpr, qn('a:ea'))
        # ea.set('typeface', 'Meiryo UI')
        if is_major:
            if text is not None:
                log(f'[{text}] Major east asian font has been converted from {ea_font_name} to +mj-ea')
            else:
                log(f'Major east asian font has been converted from {ea_font_name} to +mj-ea')
        else:
            if text is not None:
                log(f'[{text}] Minor east asian font has been converted from {ea_font_name} to +mn-ea')
            else:
                log(f'Minor east asian font has been converted from {ea_font_name} to +mn-ea')


def replace_text_frame_fonts(text_frame: TextFrame, is_major: bool):
    for paragraph in text_frame.paragraphs:
        if paragraph._element.pPr is not None and paragraph._element.pPr.defRPr is not None:
            replace_properties_fonts(paragraph._element.pPr.defRPr, is_major, None)
        for run in paragraph.runs:
            text = run.text.strip()
            replace_properties_fonts(run.font._element, is_major, text)
        if paragraph._element.endParaRPr is not None:
            replace_properties_fonts(paragraph._element.endParaRPr, is_major, None)


def replace_shape_fonts(shape: BaseShape):
    if shape.has_text_frame:
        ph = shape.element.find('.//{*}ph')
        if ph is not None and ph.get('type') == 'title':
            replace_text_frame_fonts(shape.text_frame, True)
        else:
            replace_text_frame_fonts(shape.text_frame, False)
    elif shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                replace_text_frame_fonts(cell.text_frame, False)
    elif isinstance(shape, GroupShape):
        for item in shape.shapes:
            replace_shape_fonts(item)


print(f'replace_fonts - version {version} by Shinichi Akiyama')

for file in sys.argv[1:]:
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
