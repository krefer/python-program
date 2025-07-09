"""
Генератор отчетов о проверке документов с расширенной информацией
"""
from typing import Dict, List

class ReportGenerator:
    """Генератор отчетов"""

    @staticmethod
    def print_document_validation(document_errors: List[str]):
        """Вывод ошибок документа"""
        if document_errors:
            print("\n🏗️  ОШИБКИ СТРУКТУРЫ ДОКУМЕНТА:")
            for error in document_errors:
                print(f"  ❌ {error}")
        else:
            print("\n✅ Структура документа соответствует требованиям")

    @staticmethod
    def print_progress(paragraph_num: int, classified_class: str, errors_count: int):
        """Вывод прогресса анализа"""
        print(f"\nАбзац {paragraph_num}: Классификация...")
        print(f"Определен класс: {classified_class}")
        print(f"Абзац {paragraph_num}: Проверка форматирования...")
        print(f"Абзац {paragraph_num}: Проверка содержания...")

        if errors_count > 0:
            print(f"✗ Найдено {errors_count} ошибок")
        else:
            print("✓ Ошибок не найдено")

    @staticmethod
    def print_paragraph_errors(formatting_errors: List[str], content_errors: List[str]):
        """Вывод ошибок для абзаца"""
        for error in formatting_errors:
            print(f"  📐 Форматирование: {error}")
        for error in content_errors:
            print(f"  📝 Содержание: {error}")

    @staticmethod
    def print_final_report(results: Dict):
        """Итоговый отчет о результатах анализа"""
        print("\n" + "=" * 70)
        print("ИТОГОВЫЙ ОТЧЕТ О ПРОВЕРКЕ ДОКУМЕНТА")
        print("=" * 70)

        summary = results["summary"]

        # Общая статистика
        ReportGenerator._print_summary_stats(summary)

        # Ошибки документа
        if results.get("document_errors"):
            ReportGenerator._print_document_errors(results["document_errors"])

        # Найденные классы
        ReportGenerator._print_found_classes(summary, results["paragraphs"])

        # Детализация ошибок
        ReportGenerator._print_detailed_errors(results["paragraphs"])

        # Анализ соответствия требованиям
        ReportGenerator._print_compliance_analysis(summary)

        # Рекомендации
        ReportGenerator._print_recommendations(summary)

        print("\n" + "=" * 70)

    @staticmethod
    def _print_summary_stats(summary: Dict):
        """Вывод общей статистики"""
        print(f"\n📊 ОБЩАЯ СТАТИСТИКА:")
        print(f"  • Всего абзацев: {summary['total_paragraphs']}")
        print(f"  • Всего ошибок: {summary['total_errors']}")
        print(f"  • Ошибки форматирования: {summary['formatting_errors']}")
        print(f"  • Ошибки содержания: {summary['content_errors']}")
        print(f"  • Ошибки структуры документа: {summary.get('document_errors', 0)}")

    @staticmethod
    def _print_document_errors(document_errors: List[str]):
        """Вывод ошибок структуры документа"""
        print(f"\n🏗️  ОШИБКИ СТРУКТУРЫ ДОКУМЕНТА:")
        for error in document_errors:
            print(f"  ❌ {error}")

    @staticmethod
    def _print_found_classes(summary: Dict, paragraphs: list):
        """Вывод найденных классов"""
        print(f"\n🏷️  НАЙДЕННЫЕ ЭЛЕМЕНТЫ СТАТЬИ:")

        class_translations = {
            'удк': 'УДК',
            'автор': 'Автор',
            'заголовок': 'Заголовок статьи',
            'сведения_об_авторе': 'Сведения об авторе',
            'аннотация': 'Аннотация',
            'ключевые_слова': 'Ключевые слова',
            'заголовок_английский': 'Заголовок (англ.)',
            'автор_английский': 'Автор (англ.)',
            'место_работы_английский': 'Место работы (англ.)',
            'аннотация_английская': 'Аннотация (англ.)',
            'ключевые_слова_английские': 'Ключевые слова (англ.)',
            'основной_текст': 'Основной текст'
        }

        for class_name in sorted(summary['classes_found']):
            count = sum(1 for p in paragraphs if p["classified_as"] == class_name)
            translated_name = class_translations.get(class_name, class_name)
            print(f"  • {translated_name}: {count} элемент(ов)")

    @staticmethod
    def _print_detailed_errors(paragraphs: list):
        """Вывод детальных ошибок"""
        print(f"\n🔍 ДЕТАЛИЗАЦИЯ ОШИБОК:")

        error_paragraphs = [p for p in paragraphs if p["total_errors"] > 0]

        if not error_paragraphs:
            print("  ✅ Ошибок в абзацах не найдено!")
        else:
            for para in error_paragraphs:
                print(f"\n  📄 Абзац {para['index']} ({para['classified_as']}):")
                print(f"     Текст: {para['text_preview']}")

                for error in para['formatting_errors']:
                    print(f"     ❌ Форматирование: {error}")

                for error in para['content_errors']:
                    print(f"     ❌ Содержание: {error}")

    @staticmethod
    def _print_compliance_analysis(summary: Dict):
        """Анализ соответствия требованиям"""
        print(f"\n📋 АНАЛИЗ СООТВЕТСТВИЯ ТРЕБОВАНИЯМ:")

        required_elements = [
            'удк', 'автор', 'заголовок', 'сведения_об_авторе',
            'аннотация', 'ключевые_слова', 'заголовок_английский',
            'автор_английский', 'аннотация_английская', 'ключевые_слова_английские'
        ]

        found_elements = summary['classes_found']
        missing_elements = [elem for elem in required_elements if elem not in found_elements]

        if missing_elements:
            print("  ⚠️  Отсутствующие обязательные элементы:")
            element_names = {
                'удк': 'УДК',
                'автор': 'Автор',
                'заголовок': 'Заголовок статьи',
                'сведения_об_авторе': 'Сведения об авторе',
                'аннотация': 'Аннотация',
                'ключевые_слова': 'Ключевые слова',
                'заголовок_английский': 'Заголовок на английском',
                'автор_английский': 'Автор на английском',
                'аннотация_английская': 'Аннотация на английском',
                'ключевые_слова_английские': 'Ключевые слова на английском'
            }
            for elem in missing_elements:
                print(f"    • {element_names.get(elem, elem)}")
        else:
            print("  ✅ Все обязательные элементы присутствуют")

    @staticmethod
    def _print_recommendations(summary: Dict):
        """Вывод рекомендаций"""
        print(f"\n💡 РЕКОМЕНДАЦИИ:")

        total_errors = summary['total_errors'] + summary.get('document_errors', 0)

        if total_errors == 0:
            print("  🎉 Отличная работа! Документ полностью соответствует требованиям.")
        else:
            if summary.get('document_errors', 0) > 0:
                print("  📄 Исправьте настройки документа: поля, объем, общее форматирование.")

            if summary['formatting_errors'] > summary['content_errors']:
                print("  📐 Основные проблемы в форматировании:")
                print("    - Проверьте шрифт Times New Roman для всего текста")
                print("    - Проверьте размеры шрифтов согласно требованиям")
                print("    - Проверьте выравнивание текста")
                print("    - Проверьте использование полужирного шрифта и курсива")
                print("    - Проверьте отступы абзацев (0.6 см для основного текста)")
            else:
                print("  📝 Основные проблемы в содержании:")
                print("    - Проверьте структуру и полноту всех разделов")
                print("    - Убедитесь в корректности оформления ФИО и контактов")
                print("    - Проверьте длину аннотации (300-650 символов)")
                print("    - Проверьте количество ключевых слов (4-6 слов)")

            print("  📚 Общие рекомендации:")
            print("    - Используйте только Times New Roman для всего документа")
            print("    - Соблюдайте межстрочный интервал 1.0")
            print("    - Не используйте аббревиатуры в заголовках")
            print("    - Убедитесь в корректности английской части")

