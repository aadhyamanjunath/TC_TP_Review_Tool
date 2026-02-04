"""DO-178 Test Case Review Tool - package

This package validates the 'Test Environment/Equipment' cell in each testcase row.

Public API:
- main(): CLI entry point (in main.py)
- review_env_cells(): core review logic
"""

from .logic import (
    is_empty_cell,
    load_docx_text,
    extract_section_numbers_from_docx,
    extract_section_refs,
    section_exists_in_doc,
    looks_like_instructions,
    load_equipment_keywords,
    extract_equipment,
    equipment_exists_in_doc,
    load_testcases,
    resolve_env_column,
    review_env_cells,
    save_report,
)

__all__ = [
    'is_empty_cell',
    'load_docx_text',
    'extract_section_numbers_from_docx',
    'extract_section_refs',
    'section_exists_in_doc',
    'looks_like_instructions',
    'load_equipment_keywords',
    'extract_equipment',
    'equipment_exists_in_doc',
    'load_testcases',
    'resolve_env_column',
    'review_env_cells',
    'save_report',
]
