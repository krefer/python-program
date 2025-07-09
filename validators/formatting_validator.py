"""
Валидатор форматирования текста с обновленными правилами
"""
from typing import List, Dict
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from config.criteria import FormattingCriteria

class FormattingValidator:
    """Валидатор форматирования документов"""

    @staticmethod
    def validate_document_properties(document_info: Dict) -> List[str]:
        errors = []
        requirements = FormattingCriteria.get_document_requirements()
        doc_props = document_info.get('document_properties', {})

        margins = requirements['margins']
        if doc_props.get('top_margin') and abs(doc_props['top_margin'] - margins['top'].cm) > 0.2:
            errors.append(f"Неверное верхнее поле: {doc_props['top_margin']:.1f} см (требуется 1.5 см)")

        if doc_props.get('bottom_margin') and abs(doc_props['bottom_margin'] - margins['bottom'].cm) > 0.2:
            errors.append(f"Неверное нижнее поле: {doc_props['bottom_margin']:.1f} см (требуется 1.5 см)")

        if doc_props.get('left_margin') and abs(doc_props['left_margin'] - margins['left'].cm) > 0.2:
            errors.append(f"Неверное левое поле: {doc_props['left_margin']:.1f} см (требуется 2.5 см)")

        if doc_props.get('right_margin') and abs(doc_props['right_margin'] - margins['right'].cm) > 0.2:
            errors.append(f"Неверное правое поле: {doc_props['right_margin']:.1f} см (требуется 1.0 см)")

        page_count = document_info.get('page_count', 0)
        if page_count < requirements['min_pages']:
            errors.append(f"Недостаточный объем документа: {page_count} стр. (минимум {requirements['min_pages']} стр.)")

        return errors

    @staticmethod
    def validate_formatting(para_info: Dict, expected_class: str) -> List[str]:
        errors = []
        criteria = FormattingCriteria.get_criteria(expected_class)

        if not criteria:
            return errors

        expected_font = criteria.get('font_name')
        if expected_font and para_info.get('font_name') != expected_font:
            errors.append(f"Неверный шрифт: {para_info.get('font_name')} (требуется {expected_font})")

        expected_size = criteria.get('font_size')
        actual_size = para_info.get('font_size')
        if expected_size and actual_size:
            if abs(actual_size - expected_size) > 0.2:
                errors.append(f"Неверный размер шрифта: {actual_size:.1f} (требуется {expected_size})")

        expected_alignment = criteria.get('alignment')
        actual_alignment = para_info.get('alignment')

        # Приведение к enum
        if isinstance(actual_alignment, int):
            try:
                actual_alignment = WD_ALIGN_PARAGRAPH(actual_alignment)
            except ValueError:
                actual_alignment = None

        if isinstance(expected_alignment, int):
            try:
                expected_alignment = WD_ALIGN_PARAGRAPH(expected_alignment)
            except ValueError:
                expected_alignment = None

        alignment_names = {
            WD_ALIGN_PARAGRAPH.LEFT: "по левому краю",
            WD_ALIGN_PARAGRAPH.CENTER: "по центру",
            WD_ALIGN_PARAGRAPH.RIGHT: "по правому краю",
            WD_ALIGN_PARAGRAPH.JUSTIFY: "по ширине",
            None: "не задано"
        }

        if expected_alignment is not None and actual_alignment is not None:
            if expected_alignment != actual_alignment:
                current_align = alignment_names.get(actual_alignment, "неизвестно")
                required_align = alignment_names.get(expected_alignment, "неизвестно")
                errors.append(f"Неверное выравнивание: {current_align} (требуется {required_align})")

        expected_bold = criteria.get('bold')
        if expected_bold is not None:
            if para_info.get('is_bold') != expected_bold:
                errors.append("Текст должен быть полужирным" if expected_bold else "Текст не должен быть полужирным")

        expected_italic = criteria.get('italic')
        if expected_italic is not None:
            if para_info.get('is_italic') != expected_italic:
                errors.append("Текст должен быть курсивом" if expected_italic else "Текст не должен быть курсивом")

        expected_indent = criteria.get('paragraph_indent')
        actual_indent = para_info.get('first_line_indent')
        if expected_indent and actual_indent is not None:
            if abs(actual_indent - expected_indent.cm) > 0.1:
                errors.append(f"Неверный отступ абзаца: {actual_indent:.1f} см (требуется {expected_indent.cm:.1f} см)")

        return errors
