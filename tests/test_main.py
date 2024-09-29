import pytest
from excel_to_md_table.main import excel_to_md_table

def test_excel_to_md_table():
    input_content = "Header 1\tHeader 2\nValue 1\tValue 2"
    expected_output = "| Header 1 | Header 2 |\n|---|---|\n| Value 1 | Value 2 |\n"
    assert excel_to_md_table(input_content) == expected_output