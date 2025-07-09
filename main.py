"""
Главный модуль валидатора документов с обновленными требованиями
"""
from typing import Dict
from utils.document_loader import DocumentLoader
from ai.classifier import AIClassifier, read_api_key_from_reference
from validators.formatting_validator import FormattingValidator
from validators.content_validator import ContentValidator
from reports.report_generator import ReportGenerator


api_key = read_api_key_from_reference("C:/Users/Nikita/PycharmProjects/diplom3/requirements.txt")
class DocxValidator:
    """Основной класс валидатора документов"""

    def __init__(self):
        """Инициализация компонентов"""
        self.document_loader = DocumentLoader()
        self.ai_classifier = AIClassifier(api_key=api_key)
        self.formatting_validator = FormattingValidator()
        self.content_validator = ContentValidator()
        self.report_generator = ReportGenerator()


    def analyze_document(self, file_path: str) -> Dict:
        """Полный анализ документа"""
        # Сброс состояния классификатора для нового документа
        self.ai_classifier.reset_state()
        print("Загрузка и анализ структуры документа...")

        # Загрузка документа
        document_info = self.document_loader.load_document_with_formatting(file_path)
        paragraphs_info = document_info.get('paragraphs', [])

        # Проверка общих свойств документа
        document_errors = self.formatting_validator.validate_document_properties(document_info)

        # Вывод результатов проверки документа
        self.report_generator.print_document_validation(document_errors)

        # Инициализация результатов
        results = {
            "paragraphs": [],
            "document_errors": document_errors,
            "summary": {
                "total_paragraphs": len(paragraphs_info),
                "total_errors": 0,
                "formatting_errors": 0,
                "content_errors": 0,
                "document_errors": len(document_errors),
                "classes_found": set()
            }
        }

        print("\nНачинаю анализ абзацев...")
        print("=" * 70)

        # Анализ каждого абзаца
        for i, para_info in enumerate(paragraphs_info, 1):
            paragraph_result = self._analyze_paragraph(i, para_info)

            # Сохранение результатов
            results["paragraphs"].append(paragraph_result)
            self._update_summary(results["summary"], paragraph_result)

            # Вывод прогресса
            #self.report_generator.print_progress(
            #    i,
            #    paragraph_result["classified_as"],
            ##   paragraph_result["total_errors"]
            #)

            #if paragraph_result["total_errors"] > 0:
             #   self.report_generator.print_paragraph_errors(
             #       paragraph_result["formatting_errors"],
             #       paragraph_result["content_errors"]
              #  )

        results["summary"]["classes_found"] = list(results["summary"]["classes_found"])
        return results

    def _analyze_paragraph(self, index: int, para_info: Dict) -> Dict:
        """Анализ отдельного абзаца"""
        text = para_info['text']

        # Шаг 1: Классификация
        classified_class = self.ai_classifier.classify_paragraph(
            text,
            paragraph_index=index,
            formatting_info=para_info
        )

        # Шаг 2: Проверка форматирования
        formatting_errors = self.formatting_validator.validate_formatting(para_info, classified_class)

        # Шаг 3: Проверка содержания
        content_errors = self.content_validator.validate_content(text, classified_class)

        return {
            "index": index,
            "text_preview": text[:100] + "..." if len(text) > 100 else text,
            "classified_as": classified_class,
            "formatting_errors": formatting_errors,
            "content_errors": content_errors,
            "total_errors": len(formatting_errors) + len(content_errors)
        }

    def _update_summary(self, summary: Dict, paragraph_result: Dict):
        """Обновление общей статистики"""
        summary["classes_found"].add(paragraph_result["classified_as"])
        summary["total_errors"] += paragraph_result["total_errors"]
        summary["formatting_errors"] += len(paragraph_result["formatting_errors"])
        summary["content_errors"] += len(paragraph_result["content_errors"])

    def generate_report(self, results: Dict):
        """Генерация итогового отчета"""
        self.report_generator.print_final_report(results)


if __name__ == "__main__":
    print("🚀 Инициализация валидатора документов...")
    print("Загрузка ИИ-модели для классификации текста...")

    validator = DocxValidator()

    print("\n📄 Запуск анализа документа по новым требованиям...")
    print("Проверяемые критерии:")
    print("  • Поля: верх/низ 15мм, лево 25мм, право 10мм")
    print("  • Шрифт: Times New Roman для всего текста")
    print("  • Межстрочный интервал: 1.0")
    print("  • Минимальный объем: 3 страницы")
    print("  • Форматирование всех элементов согласно ГОСТ")

    try:
        analysis_results = validator.analyze_document("test.docx")
        print("\n📊 Генерация итогового отчета...")
        validator.generate_report(analysis_results)
    except FileNotFoundError:
        print("\n❌ Файл 'test.docx' не найден!")
        print("Поместите файл для проверки в ту же папку с названием 'test.docx'")
    except Exception as e:
        print(f"\n❌ Ошибка при анализе документа: {e}")
        print("Убедитесь, что файл является корректным документом Word (.docx)")
