"""
Окно настроек критериев проверки документов
"""
import tkinter as tk
from tkinter import ttk, messagebox
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from config.criteria import FormattingCriteria
import copy


class SettingsWindow:
    """Окно настроек критериев форматирования"""

    def __init__(self, parent):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title("Настройки критериев проверки")
        self.window.geometry("800x600")
        self.window.transient(parent)
        self.window.grab_set()

        # Копируем текущие критерии
        self.current_criteria = copy.deepcopy(FormattingCriteria.CRITERIA)
        self.document_requirements = copy.deepcopy(FormattingCriteria.DOCUMENT_REQUIREMENTS)

        self.setup_ui()

    def setup_ui(self):
        """Настройка интерфейса"""
        # Создаем notebook для вкладок
        self.notebook = ttk.Notebook(self.window)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # Вкладка общих настроек документа
        self.create_document_settings_tab()

        # Вкладки для каждого типа элемента
        self.create_element_tabs()

        # Кнопки управления
        self.create_buttons()

    def create_document_settings_tab(self):
        """Создание вкладки общих настроек документа"""
        doc_frame = ttk.Frame(self.notebook)
        self.notebook.add(doc_frame, text="Документ")

        # Настройки полей
        margins_frame = ttk.LabelFrame(doc_frame, text="Поля документа (см)")
        margins_frame.pack(fill='x', padx=10, pady=5)

        self.margin_vars = {}
        margins = [('Верхнее', 'top'), ('Нижнее', 'bottom'), ('Левое', 'left'), ('Правое', 'right')]

        for i, (label, key) in enumerate(margins):
            ttk.Label(margins_frame, text=f"{label}:").grid(row=i // 2, column=(i % 2) * 2, sticky='w', padx=5, pady=2)
            var = tk.DoubleVar(value=self.document_requirements['margins'][key].cm)
            self.margin_vars[key] = var
            entry = ttk.Entry(margins_frame, textvariable=var, width=10)
            entry.grid(row=i // 2, column=(i % 2) * 2 + 1, sticky='w', padx=5, pady=2)

        # Общие настройки
        general_frame = ttk.LabelFrame(doc_frame, text="Общие настройки")
        general_frame.pack(fill='x', padx=10, pady=5)

        # Межстрочный интервал
        ttk.Label(general_frame, text="Межстрочный интервал:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.line_spacing_var = tk.DoubleVar(value=self.document_requirements['line_spacing'])
        ttk.Entry(general_frame, textvariable=self.line_spacing_var, width=10).grid(row=0, column=1, sticky='w', padx=5,
                                                                                    pady=2)

        # Шрифт по умолчанию
        ttk.Label(general_frame, text="Шрифт по умолчанию:").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.default_font_var = tk.StringVar(value=self.document_requirements['font_name'])
        font_combo = ttk.Combobox(general_frame, textvariable=self.default_font_var,
                                  values=['Times New Roman', 'Arial', 'Calibri'], width=15)
        font_combo.grid(row=1, column=1, sticky='w', padx=5, pady=2)

        # Минимальное количество страниц
        ttk.Label(general_frame, text="Мин. страниц:").grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.min_pages_var = tk.IntVar(value=self.document_requirements['min_pages'])
        ttk.Entry(general_frame, textvariable=self.min_pages_var, width=10).grid(row=2, column=1, sticky='w', padx=5,
                                                                                 pady=2)

    def create_element_tabs(self):
        """Создание вкладок для каждого типа элемента"""
        element_names = {
            'удк': 'УДК',
            'автор': 'Автор',
            'заголовок': 'Заголовок',
            'сведения_об_авторе': 'Сведения об авторе',
            'аннотация': 'Аннотация',
            'ключевые_слова': 'Ключевые слова',
            'заголовок_английский': 'Заголовок (англ)',
            'автор_английский': 'Автор (англ)',
            'место_работы_английский': 'Место работы (англ)',
            'аннотация_английская': 'Аннотация (англ)',
            'ключевые_слова_английские': 'Ключевые слова (англ)',
            'основной_текст': 'Основной текст'
        }

        self.element_vars = {}

        for element_key, element_name in element_names.items():
            if element_key in self.current_criteria:
                frame = ttk.Frame(self.notebook)
                self.notebook.add(frame, text=element_name)
                self.create_element_settings(frame, element_key)

    def create_element_settings(self, parent, element_key):
        """Создание настроек для конкретного элемента"""
        criteria = self.current_criteria[element_key]

        # Словарь для хранения переменных этого элемента
        self.element_vars[element_key] = {}

        # Фрейм форматирования
        format_frame = ttk.LabelFrame(parent, text="Форматирование")
        format_frame.pack(fill='x', padx=10, pady=5)

        row = 0

        # Шрифт
        if 'font_name' in criteria:
            ttk.Label(format_frame, text="Шрифт:").grid(row=row, column=0, sticky='w', padx=5, pady=2)
            var = tk.StringVar(value=criteria['font_name'])
            self.element_vars[element_key]['font_name'] = var
            combo = ttk.Combobox(format_frame, textvariable=var,
                                 values=['Times New Roman', 'Arial', 'Calibri'], width=15)
            combo.grid(row=row, column=1, sticky='w', padx=5, pady=2)
            row += 1

        # Размер шрифта
        if 'font_size' in criteria:
            ttk.Label(format_frame, text="Размер шрифта:").grid(row=row, column=0, sticky='w', padx=5, pady=2)
            var = tk.DoubleVar(value=criteria['font_size'])
            self.element_vars[element_key]['font_size'] = var
            ttk.Entry(format_frame, textvariable=var, width=10).grid(row=row, column=1, sticky='w', padx=5, pady=2)
            row += 1

        # Выравнивание
        if 'alignment' in criteria:
            ttk.Label(format_frame, text="Выравнивание:").grid(row=row, column=0, sticky='w', padx=5, pady=2)

            alignment_map = {
                WD_ALIGN_PARAGRAPH.LEFT: "По левому краю",
                WD_ALIGN_PARAGRAPH.CENTER: "По центру",
                WD_ALIGN_PARAGRAPH.RIGHT: "По правому краю",
                WD_ALIGN_PARAGRAPH.JUSTIFY: "По ширине"
            }

            current_alignment = criteria['alignment']
            current_text = alignment_map.get(current_alignment, "По левому краю")

            var = tk.StringVar(value=current_text)
            self.element_vars[element_key]['alignment'] = var
            combo = ttk.Combobox(format_frame, textvariable=var,
                                 values=list(alignment_map.values()), width=15)
            combo.grid(row=row, column=1, sticky='w', padx=5, pady=2)
            row += 1

        # Полужирный
        if 'bold' in criteria:
            var = tk.BooleanVar(value=criteria['bold'])
            self.element_vars[element_key]['bold'] = var
            ttk.Checkbutton(format_frame, text="Полужирный", variable=var).grid(row=row, column=0, columnspan=2,
                                                                                sticky='w', padx=5, pady=2)
            row += 1

        # Курсив
        if 'italic' in criteria:
            var = tk.BooleanVar(value=criteria['italic'])
            self.element_vars[element_key]['italic'] = var
            ttk.Checkbutton(format_frame, text="Курсив", variable=var).grid(row=row, column=0, columnspan=2, sticky='w',
                                                                            padx=5, pady=2)
            row += 1

        # Отступ абзаца
        if 'paragraph_indent' in criteria:
            ttk.Label(format_frame, text="Отступ абзаца (см):").grid(row=row, column=0, sticky='w', padx=5, pady=2)
            var = tk.DoubleVar(value=criteria['paragraph_indent'].cm)
            self.element_vars[element_key]['paragraph_indent'] = var
            ttk.Entry(format_frame, textvariable=var, width=10).grid(row=row, column=1, sticky='w', padx=5, pady=2)
            row += 1

        # Правила содержания
        if 'content_rules' in criteria and criteria['content_rules']:
            content_frame = ttk.LabelFrame(parent, text="Правила содержания")
            content_frame.pack(fill='both', expand=True, padx=10, pady=5)

            # Создаем текстовое поле для отображения правил
            text_widget = tk.Text(content_frame, height=8, wrap='word')
            scrollbar = ttk.Scrollbar(content_frame, orient="vertical", command=text_widget.yview)
            text_widget.configure(yscrollcommand=scrollbar.set)

            text_widget.pack(side='left', fill='both', expand=True)
            scrollbar.pack(side='right', fill='y')

            # Добавляем правила в текстовое поле
            rules_text = "Активные правила проверки содержания:\n\n"
            for i, (rule_name, _) in enumerate(criteria['content_rules'], 1):
                rules_text += f"{i}. {rule_name}\n"

            text_widget.insert('1.0', rules_text)
            text_widget.config(state='disabled')  # Только для чтения

    def create_buttons(self):
        """Создание кнопок управления"""
        button_frame = ttk.Frame(self.window)
        button_frame.pack(fill='x', padx=10, pady=10)

        ttk.Button(button_frame, text="Сохранить", command=self.save_settings).pack(side='right', padx=5)
        ttk.Button(button_frame, text="Отмена", command=self.window.destroy).pack(side='right', padx=5)
        ttk.Button(button_frame, text="Сбросить", command=self.reset_settings).pack(side='left', padx=5)
        ttk.Button(button_frame, text="По умолчанию", command=self.load_defaults).pack(side='left', padx=5)

    def save_settings(self):
        """Сохранение настроек"""
        try:
            # Обновляем требования к документу
            self.document_requirements['margins']['top'] = Cm(self.margin_vars['top'].get())
            self.document_requirements['margins']['bottom'] = Cm(self.margin_vars['bottom'].get())
            self.document_requirements['margins']['left'] = Cm(self.margin_vars['left'].get())
            self.document_requirements['margins']['right'] = Cm(self.margin_vars['right'].get())
            self.document_requirements['line_spacing'] = self.line_spacing_var.get()
            self.document_requirements['font_name'] = self.default_font_var.get()
            self.document_requirements['min_pages'] = self.min_pages_var.get()

            # Обновляем критерии элементов
            alignment_map_reverse = {
                "По левому краю": WD_ALIGN_PARAGRAPH.LEFT,
                "По центру": WD_ALIGN_PARAGRAPH.CENTER,
                "По правому краю": WD_ALIGN_PARAGRAPH.RIGHT,
                "По ширине": WD_ALIGN_PARAGRAPH.JUSTIFY
            }

            for element_key, vars_dict in self.element_vars.items():
                criteria = self.current_criteria[element_key]

                for var_name, var in vars_dict.items():
                    if var_name == 'alignment':
                        alignment_text = var.get()
                        criteria[var_name] = alignment_map_reverse.get(alignment_text, WD_ALIGN_PARAGRAPH.LEFT)
                    elif var_name == 'paragraph_indent':
                        criteria[var_name] = Cm(var.get())
                    else:
                        criteria[var_name] = var.get()

            # Применяем изменения к основному классу
            FormattingCriteria.CRITERIA = self.current_criteria
            FormattingCriteria.DOCUMENT_REQUIREMENTS = self.document_requirements

            messagebox.showinfo("Настройки", "Настройки успешно сохранены!")
            self.window.destroy()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при сохранении настроек: {str(e)}")

    def reset_settings(self):
        """Сброс настроек к текущим значениям"""
        if messagebox.askyesno("Сброс", "Сбросить все изменения к текущим настройкам?"):
            self.window.destroy()
            SettingsWindow(self.parent)

    def load_defaults(self):
        """Загрузка настроек по умолчанию"""
        if messagebox.askyesno("По умолчанию",
                               "Загрузить настройки по умолчанию? Все текущие изменения будут потеряны."):
            # Перезагружаем класс критериев
            import importlib
            from config import criteria
            importlib.reload(criteria)

            self.current_criteria = copy.deepcopy(criteria.FormattingCriteria.CRITERIA)
            self.document_requirements = copy.deepcopy(criteria.FormattingCriteria.DOCUMENT_REQUIREMENTS)

            self.window.destroy()
            SettingsWindow(self.parent)