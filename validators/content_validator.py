"""
Валидатор содержания текста с обновленными правилами
"""
from typing import List
from config.criteria import FormattingCriteria

class ContentValidator:
    """Валидатор содержания документов"""

    @staticmethod
    def validate_content(text: str, expected_class: str) -> List[str]:
        """Проверка содержания согласно критериям"""
        errors = []
        criteria = FormattingCriteria.get_criteria(expected_class)

        if not criteria:
            return errors

        content_rules = criteria.get('content_rules', [])

        for rule_name, rule_func in content_rules:
            try:
                if not rule_func(text):
                    errors.append(f"{rule_name}")
            except Exception as e:
                errors.append(f"Ошибка проверки правила '{rule_name}': {str(e)}")

        return errors