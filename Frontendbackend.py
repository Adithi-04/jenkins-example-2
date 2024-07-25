# test_core_logic.py
import pytest
import re

def test_extract_font_details(rtf_content):
    # Extract font table
    font_table_pattern = re.compile(r'{\\fonttbl(.*)}', re.DOTALL)
    font_table_match = font_table_pattern.search(rtf_content)
    if font_table_match:
        font_table = font_table_match.group(1)
    else:
        return {}

    # Extract font details from font table
    font_pattern = re.compile(r'{\\f(\d+)\\.*? ([^;]+?);}')
    fonts = {}
    for match in font_pattern.finditer(font_table):
        font_id, font_name = match.groups()
        fonts['f'+font_id] = font_name

    return fonts

