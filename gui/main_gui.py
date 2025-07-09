"""
Главное окно GUI для валидатора документов Word
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import sys
from typing import Dict, List

# Добавляем родительскую папку в путь для импорта модулей
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import DocxValidator
from gui.settings_window import SettingsWindow
from config.criteria import FormattingCriteria


class ValidatorGUI:
    """Главное окно GUI валидатора документов"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Валидатор документов Word - Проверка форматирования")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)

        # Настройка стиля
        style = ttk.Style()
        style.theme_use('clam')

        # Инициализация валидатора
        self.validator = None
        self.current_file_path = None
        self.analysis_results = None

        # Создание интерфейса
        self.create_widgets()
        self.setup_layout()

        # Инициализация валидатора в отдельном потоке
        self.init_validator_async()

    def create_widgets(self):
        """Создание виджетов интерфейса"""
        # Главное меню
        self.create_menu()

        # Фрейм для выбора файла
        self.file_frame = ttk.LabelFrame(self.root, text="Выбор документа", padding=10)

        self.file_path_var = tk.StringVar()
        self.file_entry = ttk.Entry(self.file_frame, textvariable=self.file_path_var,
                                    state='readonly', width=60)
        self.browse_button = ttk.Button(self.file_frame, text="Обзор...",
                                        command=self.browse_file)

        # Фрейм для кнопок управления
        self.control_frame = ttk.Frame(self.root)

        self.analyze_button = ttk.Button(self.control_frame, text="Анализировать документ",
                                         command=self.start_analysis, state='disabled')
        self.settings_button = ttk.Button(self.control_frame, text="Настройки критериев",
                                          command=self.open_settings)
        self.clear_button = ttk.Button(self.control_frame, text="Очистить результаты",
                                       command=self.clear_results)

        # Прогресс-бар
        self.progress_var = tk.StringVar(value="Готов к работе")
        self.progress_label = ttk.Label(self.root, textvariable=self.progress_var)
        self.progress_bar = ttk.Progressbar(self.root, mode='indeterminate')

        # Notebook для результатов
        self.notebook = ttk.Notebook(self.root)

        # Вкладка "Общий отчет"
        self.summary_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.summary_frame, text="Общий отчет")

        self.summary_text = scrolledtext.ScrolledText(self.summary_frame, wrap=tk.WORD,
                                                      height=15, font=('Consolas', 10))

        # Вкладка "Детальный анализ"
        self.detail_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.detail_frame, text="Детальный анализ")

        # Фрейм для фильтров в детальном анализе
        self.filter_frame = ttk.LabelFrame(self.detail_frame, text="Фильтры", padding=5)

        self.show_errors_only = tk.BooleanVar(value=False)
        self.errors_checkbox = ttk.Checkbutton(self.filter_frame, text="Показать только ошибки",
                                               variable=self.show_errors_only,
                                               command=self.update_detail_view)

        self.class_filter_var = tk.StringVar(value="Все")
        self.class_filter_label = ttk.Label(self.filter_frame, text="Класс:")
        self.class_filter_combo = ttk.Combobox(self.filter_frame, textvariable=self.class_filter_var,
                                               values=["Все"], state="readonly")
        self.class_filter_combo.bind('<<ComboboxSelected>>', lambda e: self.update_detail_view())

        # Дерево для детального анализа
        self.detail_tree = ttk.Treeview(self.detail_frame, columns=('index', 'class', 'errors', 'text'),
                                        show='tree headings', height=12)

        self.detail_tree.heading('#0', text='Тип')
        self.detail_tree.heading('index', text='№')
        self.detail_tree.heading('class', text='Класс')
        self.detail_tree.heading('errors', text='Ошибки')
        self.detail_tree.heading('text', text='Текст (первые 50 символов)')

        self.detail_tree.column('#0', width=30)
        self.detail_tree.column('index', width=50)
        self.detail_tree.column('class', width=150)
        self.detail_tree.column('errors', width=80)
        self.detail_tree.column('text', width=400)

        # Скроллбар для дерева
        self.detail_scrollbar = ttk.Scrollbar(self.detail_frame, orient=tk.VERTICAL,
                                              command=self.detail_tree.yview)
        self.detail_tree.configure(yscrollcommand=self.detail_scrollbar.set)

        # Статус-бар
        self.status_frame = ttk.Frame(self.root)
        self.status_var = tk.StringVar(value="Инициализация...")
        self.status_label = ttk.Label(self.status_frame, textvariable=self.status_var,
                                      relief=tk.SUNKEN, anchor=tk.W)

    def create_menu(self):
        """Создание главного меню"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # Меню "Файл"
        self.file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Файл", menu=self.file_menu)
        self.file_menu.add_command(label="Открыть документ...", command=self.browse_file, accelerator="Ctrl+O")
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Сохранить отчет...", command=self.save_report, state='disabled')
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Выход", command=self.root.quit)

        # Меню "Анализ"
        analysis_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Анализ", menu=analysis_menu)
        analysis_menu.add_command(label="Запустить анализ", command=self.start_analysis, accelerator="F5")
        analysis_menu.add_command(label="Остановить анализ", command=self.stop_analysis, state='disabled')
        analysis_menu.add_separator()
        analysis_menu.add_command(label="Настройки критериев", command=self.open_settings)

        # Меню "Вид"
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Вид", menu=view_menu)
        view_menu.add_command(label="Очистить результаты", command=self.clear_results)

        # Меню "Справка"
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Справка", menu=help_menu)
        help_menu.add_command(label="О программе", command=self.show_about)

        # Привязка горячих клавиш
        self.root.bind('<Control-o>', lambda e: self.browse_file())
        self.root.bind('<F5>', lambda e: self.start_analysis())

    def setup_layout(self):
        """Размещение виджетов"""
        # Фрейм выбора файла
        self.file_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        self.browse_button.pack(side=tk.RIGHT)

        # Фрейм управления
        self.control_frame.pack(fill=tk.X, padx=10, pady=5)
        self.analyze_button.pack(side=tk.LEFT, padx=(0, 10))
        self.settings_button.pack(side=tk.LEFT, padx=(0, 10))
        self.clear_button.pack(side=tk.LEFT)

        # Прогресс
        self.progress_label.pack(fill=tk.X, padx=10, pady=(10, 2))
        self.progress_bar.pack(fill=tk.X, padx=10, pady=(0, 10))

        # Notebook с результатами
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # Вкладка общего отчета
        self.summary_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Вкладка детального анализа
        self.filter_frame.pack(fill=tk.X, padx=5, pady=(5, 0))
        self.errors_checkbox.pack(side=tk.LEFT, padx=(0, 20))
        self.class_filter_label.pack(side=tk.LEFT, padx=(0, 5))
        self.class_filter_combo.pack(side=tk.LEFT)

        self.detail_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0), pady=5)
        self.detail_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 5), pady=5)

        # Статус-бар
        self.status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        self.status_label.pack(fill=tk.X, padx=2, pady=2)

    def init_validator_async(self):
        """Асинхронная инициализация валидатора"""

        def init_worker():
            try:
                self.validator = DocxValidator()
                self.root.after(0, self.on_validator_ready)
            except Exception as e:
                self.root.after(0, self.on_validator_error, str(e))

        thread = threading.Thread(target=init_worker, daemon=True)
        thread.start()

    def on_validator_ready(self):
        """Обработчик готовности валидатора"""
        self.status_var.set("Готов к работе")
        self.analyze_button.config(state='normal')

    def on_validator_error(self, error_msg):
        """Обработчик ошибки инициализации валидатора"""
        self.status_var.set(f"Ошибка инициализации: {error_msg}")
        messagebox.showerror("Ошибка", f"Не удалось инициализировать валидатор:\n{error_msg}")

    def browse_file(self):
        """Выбор файла для анализа"""
        file_path = filedialog.askopenfilename(
            title="Выберите документ Word",
            filetypes=[
                ("Документы Word", "*.docx"),
                ("Все файлы", "*.*")
            ]
        )

        if file_path:
            self.current_file_path = file_path
            self.file_path_var.set(file_path)
            self.status_var.set(f"Выбран файл: {os.path.basename(file_path)}")

    def start_analysis(self):
        """Запуск анализа документа"""
        if not self.current_file_path:
            messagebox.showwarning("Предупреждение", "Сначала выберите документ для анализа!")
            return

        if not self.validator:
            messagebox.showerror("Ошибка", "Валидатор не инициализирован!")
            return

        # Блокировка интерфейса
        self.analyze_button.config(state='disabled')
        self.browse_button.config(state='disabled')
        self.progress_bar.start()
        self.progress_var.set("Анализ документа...")

        # Запуск анализа в отдельном потоке
        def analysis_worker():
            try:
                results = self.validator.analyze_document(self.current_file_path)
                self.root.after(0, self.on_analysis_complete, results)
            except Exception as e:
                self.root.after(0, self.on_analysis_error, str(e))

        thread = threading.Thread(target=analysis_worker, daemon=True)
        thread.start()

    def on_analysis_complete(self, results):
        """Обработчик завершения анализа"""
        self.analysis_results = results

        # Разблокировка интерфейса
        self.analyze_button.config(state='normal')
        self.browse_button.config(state='normal')
        self.progress_bar.stop()
        self.file_menu.entryconfig("Сохранить отчет...", state='normal')
        self.progress_var.set("Анализ завершен")

        # Обновление результатов
        self.update_summary_view()
        self.update_detail_view()

        # Обновление статуса
        summary = results['summary']
        total_errors = summary['total_errors'] + summary.get('document_errors', 0)
        self.status_var.set(f"Анализ завершен. Найдено ошибок: {total_errors}")

        # Включение сохранения отчета
        self.root.nametowidget('.!menu').entryconfig("Файл", state='normal')

    def on_analysis_error(self, error_msg):
        """Обработчик ошибки анализа"""
        # Разблокировка интерфейса
        self.analyze_button.config(state='normal')
        self.browse_button.config(state='normal')
        self.progress_bar.stop()
        self.progress_var.set("Ошибка анализа")

        messagebox.showerror("Ошибка анализа", f"Произошла ошибка при анализе документа:\n{error_msg}")
        self.status_var.set(f"Ошибка: {error_msg}")

    def update_summary_view(self):
        """Обновление общего отчета"""
        if not self.analysis_results:
            return

        # Очистка текста
        self.summary_text.delete(1.0, tk.END)

        # Генерация отчета
        report_text = self.generate_summary_report()
        self.summary_text.insert(1.0, report_text)

        # Настройка цветов для ошибок
        self.highlight_errors_in_summary()

    def generate_summary_report(self):
        """Генерация текста общего отчета"""
        results = self.analysis_results
        summary = results['summary']

        report_lines = []
        report_lines.append("=" * 70)
        report_lines.append("ОТЧЕТ О ПРОВЕРКЕ ДОКУМЕНТА")
        report_lines.append("=" * 70)
        report_lines.append("")

        # Общая статистика
        report_lines.append("📊 ОБЩАЯ СТАТИСТИКА:")
        report_lines.append(f"  • Всего абзацев: {summary['total_paragraphs']}")
        report_lines.append(f"  • Всего ошибок: {summary['total_errors']}")
        report_lines.append(f"  • Ошибки форматирования: {summary['formatting_errors']}")
        report_lines.append(f"  • Ошибки содержания: {summary['content_errors']}")
        report_lines.append(f"  • Ошибки структуры документа: {summary.get('document_errors', 0)}")
        report_lines.append("")

        # Ошибки документа
        if results.get('document_errors'):
            report_lines.append("🏗️ ОШИБКИ СТРУКТУРЫ ДОКУМЕНТА:")
            for error in results['document_errors']:
                report_lines.append(f"  ❌ {error}")
            report_lines.append("")

        # Найденные классы
        report_lines.append("🏷️ НАЙДЕННЫЕ ЭЛЕМЕНТЫ СТАТЬИ:")
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
            count = sum(1 for p in results['paragraphs'] if p["classified_as"] == class_name)
            translated_name = class_translations.get(class_name, class_name)
            report_lines.append(f"  • {translated_name}: {count} элемент(ов)")
        report_lines.append("")

        # Анализ соответствия
        report_lines.append("📋 АНАЛИЗ СООТВЕТСТВИЯ ТРЕБОВАНИЯМ:")
        required_elements = [
            'удк', 'автор', 'заголовок', 'сведения_об_авторе',
            'аннотация', 'ключевые_слова', 'заголовок_английский',
            'автор_английский', 'аннотация_английская', 'ключевые_слова_английские'
        ]

        found_elements = summary['classes_found']
        missing_elements = [elem for elem in required_elements if elem not in found_elements]

        if missing_elements:
            report_lines.append("  ⚠️ Отсутствующие обязательные элементы:")
            for elem in missing_elements:
                translated_name = class_translations.get(elem, elem)
                report_lines.append(f"    • {translated_name}")
        else:
            report_lines.append("  ✅ Все обязательные элементы присутствуют")

        report_lines.append("")
        report_lines.append("=" * 70)

        return "\n".join(report_lines)

    def highlight_errors_in_summary(self):
        """Выделение ошибок в общем отчете"""
        # Настройка тегов для выделения
        self.summary_text.tag_configure("error", foreground="red")
        self.summary_text.tag_configure("warning", foreground="orange")
        self.summary_text.tag_configure("success", foreground="green")

        # Поиск и выделение ошибок
        content = self.summary_text.get(1.0, tk.END)
        lines = content.split('\n')

        for i, line in enumerate(lines):
            if '❌' in line or 'ОШИБКИ' in line:
                start = f"{i + 1}.0"
                end = f"{i + 1}.end"
                self.summary_text.tag_add("error", start, end)
            elif '⚠️' in line or 'Отсутствующие' in line:
                start = f"{i + 1}.0"
                end = f"{i + 1}.end"
                self.summary_text.tag_add("warning", start, end)
            elif '✅' in line:
                start = f"{i + 1}.0"
                end = f"{i + 1}.end"
                self.summary_text.tag_add("success", start, end)

    def update_detail_view(self):
        """Обновление детального анализа"""
        if not self.analysis_results:
            return

        # Очистка дерева
        for item in self.detail_tree.get_children():
            self.detail_tree.delete(item)

        # Обновление фильтра классов
        classes = set()
        for para in self.analysis_results['paragraphs']:
            classes.add(para['classified_as'])

        class_values = ["Все"] + sorted(list(classes))
        self.class_filter_combo['values'] = class_values

        # Фильтрация параграфов
        paragraphs = self.analysis_results['paragraphs']

        show_errors_only = self.show_errors_only.get()
        class_filter = self.class_filter_var.get()

        filtered_paragraphs = []
        for para in paragraphs:
            # Фильтр по ошибкам
            if show_errors_only and para['total_errors'] == 0:
                continue

            # Фильтр по классу
            if class_filter != "Все" and para['classified_as'] != class_filter:
                continue

            filtered_paragraphs.append(para)

        # Заполнение дерева
        for para in filtered_paragraphs:
            # Главная запись параграфа
            icon = "📄" if para['total_errors'] == 0 else "❌"
            error_text = f"{para['total_errors']}" if para['total_errors'] > 0 else "Нет"

            parent_id = self.detail_tree.insert('', 'end',
                                                text=icon,
                                                values=(
                                                para['index'], para['classified_as'], error_text, para['text_preview']),
                                                tags=('error' if para['total_errors'] > 0 else 'normal',)
                                                )

            # Добавление ошибок как дочерних элементов
            if para['total_errors'] > 0:
                for error in para['formatting_errors']:
                    self.detail_tree.insert(parent_id, 'end',
                                            text="📐",
                                            values=("", "Форматирование", "", error),
                                            tags=('formatting_error',)
                                            )

                for error in para['content_errors']:
                    self.detail_tree.insert(parent_id, 'end',
                                            text="📝",
                                            values=("", "Содержание", "", error),
                                            tags=('content_error',)
                                            )

        # Настройка цветов
        self.detail_tree.tag_configure('error', foreground='red')
        self.detail_tree.tag_configure('formatting_error', foreground='blue')
        self.detail_tree.tag_configure('content_error', foreground='purple')
        self.detail_tree.tag_configure('normal', foreground='black')

    def clear_results(self):
        """Очистка результатов анализа"""
        self.analysis_results = None
        self.summary_text.delete(1.0, tk.END)

        for item in self.detail_tree.get_children():
            self.detail_tree.delete(item)

        self.class_filter_combo['values'] = ["Все"]
        self.class_filter_var.set("Все")
        self.show_errors_only.set(False)

        self.status_var.set("Результаты очищены")

    def open_settings(self):
        """Открытие окна настроек"""
        settings_window = SettingsWindow(self.root)

    def save_report(self):
        """Сохранение отчета в файл"""
        if not self.analysis_results:
            messagebox.showwarning("Предупреждение", "Нет результатов для сохранения!")
            return

        file_path = filedialog.asksaveasfilename(
            title="Сохранить отчет",
            defaultextension=".txt",
            filetypes=[
                ("Текстовые файлы", "*.txt"),
                ("Все файлы", "*.*")
            ]
        )

        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.generate_summary_report())
                    f.write("\n\nДЕТАЛЬНЫЙ АНАЛИЗ ОШИБОК:\n")
                    f.write("=" * 50 + "\n")

                    for para in self.analysis_results['paragraphs']:
                        if para['total_errors'] > 0:
                            f.write(f"\nАбзац {para['index']} ({para['classified_as']}):\n")
                            f.write(f"Текст: {para['text_preview']}\n")

                            for error in para['formatting_errors']:
                                f.write(f"  📐 Форматирование: {error}\n")

                            for error in para['content_errors']:
                                f.write(f"  📝 Содержание: {error}\n")

                messagebox.showinfo("Успех", f"Отчет сохранен в файл:\n{file_path}")

            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить отчет:\n{str(e)}")

    def stop_analysis(self):
        """Остановка анализа (заглушка)"""
        messagebox.showinfo("Информация", "Функция остановки анализа пока не реализована")

    def show_about(self):
        """Показ информации о программе"""
        about_text = """
Валидатор документов Word
Версия 1.0

Программа для проверки форматирования научных статей
в соответствии с требованиями ГОСТ.

Возможности:
• Автоматическая классификация элементов статьи
• Проверка форматирования (шрифт, размер, выравнивание)
• Проверка содержания согласно критериям
• Настройка параметров проверки
• Детальные отчеты об ошибках

© 2025 Validator Team
"""
        messagebox.showinfo("О программе", about_text)

    def run(self):
        """Запуск главного цикла приложения"""
        self.root.mainloop()


def main():
    """Главная функция запуска GUI"""
    app = ValidatorGUI()
    app.run()


if __name__ == "__main__":
    main()