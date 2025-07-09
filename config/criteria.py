"""
Конфигурация критериев для проверки документов по новым требованиям
"""
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm,Mm
import re
from typing import Dict, List

class FormattingCriteria:
    """Критерии форматирования для различных типов текста"""

    # Общие требования к документу
    DOCUMENT_REQUIREMENTS = {
        "margins": {
            "top": Mm(15),  # 15 мм
            "bottom": Mm(15),  # 15 мм
            "left": Mm(25),  # 25 мм
            "right": Mm(10)  # 10 мм
        },
        "line_spacing": 1.0,
        "font_name": "Times New Roman",
        "min_pages": 3
    }

    CRITERIA = {
        "удк": {
            "font_name": "Times New Roman",
            "font_size": 10.5,
            "alignment": WD_ALIGN_PARAGRAPH.LEFT,
            "bold": False,
            "italic": False,
            "content_rules": [
                ("должен начинаться с 'УДК'", lambda t: t.strip().upper().startswith('УДК')),
                ("должен содержать код классификации", lambda t: len(t.split()) >= 2)
            ]
        },
        "автор": {
            "font_name": "Times New Roman",
            "font_size": 12,
            "alignment": WD_ALIGN_PARAGRAPH.LEFT,
            "bold": False,
            "italic": False,
            "content_rules": [
                ("должен содержать корректный формат ФИО с инициалами", lambda t: FormattingCriteria._check_author_format_improved(t)),
                ("инициалы без пробелов", lambda t: FormattingCriteria._check_initials_format(t))
            ]
        },
        "заголовок": {
            "font_name": "Times New Roman",
            "font_size": 12,
            "alignment": WD_ALIGN_PARAGRAPH.LEFT,
            "bold": True,
            "italic": False,
            "content_rules": [
                ("не должен содержать аббревиатуры", lambda t: not FormattingCriteria._has_abbreviations(t)),
                ("должен начинаться с заглавной буквы", lambda t: t[0].isupper() if t else False)
            ]
        },
        "сведения_об_авторе": {
            "font_name": "Times New Roman",
            "font_size": 10.5,
            "alignment": WD_ALIGN_PARAGRAPH.LEFT,
            "bold": False,
            "italic": False,
            "content_rules": [
                ("должен содержать профессиональную информацию", lambda t: FormattingCriteria._has_professional_info(t)),
                ("должен быть в именительном падеже", lambda t: True)  # Упрощенная проверка
            ]
        },
        "аннотация": {
            "font_name": "Times New Roman",
            "font_size": 10.5,
            "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "bold": False,
            "italic": True,
            "content_rules": [
                ("должна быть 300-650 символов", lambda t: 300 <= len(t) <= 650),
                ("не должна содержать заголовок 'Аннотация'", lambda t: 'аннотация' not in t.lower()[:20])
            ]
        },
        "ключевые_слова": {
            "font_name": "Times New Roman",
            "font_size": 10.5,
            "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "bold": False,
            "italic": True,
            "content_rules": [
                ("должно быть 4-6 ключевых слов", lambda t: 4 <= len([w.strip() for w in t.split(',') if w.strip()]) <= 6),
                ("не более 100 символов", lambda t: len(t) <= 100)
            ]
        },
        "заголовок_английский": {
            "font_name": "Times New Roman",
            "font_size": 10.5,
            "alignment": WD_ALIGN_PARAGRAPH.LEFT,
            "bold": True,
            "italic": False,
            "content_rules": [
                ("должен содержать английский текст", lambda t: re.search(r'[a-zA-Z]', t)),
                ("должен быть корректно написан", lambda t: len(re.findall(r'[a-zA-Z]+', t)) >= 3)
            ]
        },
        "автор_английский": {
            "font_name": "Times New Roman",
            "font_size": 10.5,
            "alignment": WD_ALIGN_PARAGRAPH.LEFT,
            "bold": False,
            "italic": False,
            "content_rules": [
                ("фамилия + инициалы без пробела", lambda t: FormattingCriteria._check_english_author_format(t)),
                ("должен содержать английский текст", lambda t: re.search(r'[a-zA-Z]', t))
            ]
        },
        "место_работы_английский": {
            "font_name": "Times New Roman",
            "font_size": 10.5,
            "alignment": WD_ALIGN_PARAGRAPH.LEFT,
            "bold": False,
            "italic": False,
            "content_rules": [
                ("должно содержать название организации", lambda t: len(t.split()) >= 3),
                ("должно содержать город и страну", lambda t: ',' in t)
            ]
        },
        "аннотация_английская": {
            "font_name": "Times New Roman",
            "font_size": 10.5,
            "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "bold": False,
            "italic": True,
            "content_rules": [
                ("должна содержать английский текст", lambda t: re.search(r'[a-zA-Z]', t)),
                ("должна быть достаточной длины", lambda t: len(t) >= 100)
            ]
        },
        "ключевые_слова_английские": {
            "font_name": "Times New Roman",
            "font_size": 10.5,
            "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "bold": False,
            "italic": True,
            "content_rules": [
                ("должны содержать английский текст", lambda t: re.search(r'[a-zA-Z]', t)),
                ("должно быть 4-6 слов", lambda t: 4 <= len([w.strip() for w in t.split(',') if w.strip()]) <= 6)
            ]
        },
        "основной_текст": {
            "font_name": "Times New Roman",
            "font_size": 10.5,
            "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "bold": False,
            "italic": False,
            "paragraph_indent": Cm(0.6),  # 0.6 см отступ абзаца
            "content_rules": [
                ("должен содержать законченные предложения", lambda t: t.count('.') >= 1),
                ("не должен быть слишком коротким", lambda t: len(t.split()) >= 10)
            ]
        }
    }

    @staticmethod
    def _has_abbreviations(text: str) -> bool:
        """Проверка на наличие аббревиатур"""
        # Исключаем из проверки известные ученые степени и сокращения
        allowed_abbr = ['к.т.н', 'д.т.н', 'к.э.н', 'д.э.н', 'к.ф.-м.н', 'д.ф.-м.н']

        # Удаляем разрешенные сокращения для проверки
        text_clean = text
        for abbr in allowed_abbr:
            text_clean = text_clean.replace(abbr, '')

        # Ищем аббревиатуры (2-6 заглавных букв подряд)
        abbreviations = re.findall(r'\b[А-ЯA-Z]{2,6}\b', text_clean)
        return len(abbreviations) > 0

    @staticmethod
    def _check_author_format_improved(text: str) -> bool:
        """Улучшенная проверка формата авторов - может быть несколько через запятую"""
        # Убираем лишние пробелы и разделяем по запятым
        authors = [author.strip() for author in text.split(',')]

        # Проверяем каждого автора
        for author in authors:
            if not author:
                continue
            # Паттерн: Фамилия И.О. (с возможными пробелами)
            pattern = r'^[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.$'
            if not re.match(pattern, author.strip()):
                return False

        return len(authors) > 0

    @staticmethod
    def _check_initials_format(text: str) -> bool:
        """Проверка формата инициалов без пробелов"""
        if '.' in text:
            # Ищем инициалы в тексте
            initials_matches = re.findall(r'[А-ЯЁ]\.[А-ЯЁ]\.', text)
            if initials_matches:
                # Проверяем, что между инициалами нет пробелов
                for match in initials_matches:
                    if ' ' in match:
                        return False
        return True

    @staticmethod
    def _check_english_author_format(text: str) -> bool:
        """Проверка формата английского автора"""
        pattern = r'[A-Z][a-z]+\s*[A-Z]\.[A-Z]\.'
        return bool(re.search(pattern, text))

    @staticmethod
    def _has_full_name_complete(text: str) -> bool:
        """Проверка на наличие полного ФИО"""
        # Проверяем наличие имени, отчества и фамилии
        words = text.split()
        return len(words) >= 3 and any(word for word in words if len(word) > 3)

    @staticmethod
    def _has_professional_info(text: str) -> bool:
        """Проверка профессиональной информации"""
        info_keywords = [
            'к.т.н', 'д.т.н', 'кандидат', 'доктор', 'профессор', 'доцент',
            'аспирант', 'магистр', 'заведующий', 'директор', 'кафедра',
            'университет', 'институт', 'факультет', 'область', 'город'
        ]
        text_lower = text.lower()
        return any(keyword in text_lower for keyword in info_keywords)

    @staticmethod
    def _has_workplace_info(text: str) -> bool:
        """Проверка информации о месте работы"""
        workplace_keywords = ['университет', 'институт', 'академия', 'центр', 'кафедра', 'факультет']
        text_lower = text.lower()
        return any(keyword in text_lower for keyword in workplace_keywords)

    @classmethod
    def get_criteria(cls, class_name: str):
        """Получить критерии для определенного класса"""
        return cls.CRITERIA.get(class_name, {})

    @classmethod
    def get_all_classes(cls):
        """Получить список всех доступных классов"""
        return list(cls.CRITERIA.keys())

    @classmethod
    def get_document_requirements(cls):
        """Получить общие требования к документу"""
        return cls.DOCUMENT_REQUIREMENTS