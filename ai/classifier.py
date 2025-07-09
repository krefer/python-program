"""
Улучшенный ИИ-классификатор для определения типов текста с учетом валидации документов
"""
import requests
import json
import re
from time import sleep
from typing import List, Dict, Optional
from langdetect import detect

def read_api_key_from_reference(file_path="requirements.txt"):
    """
    Читает API ключ из файла Reference.txt.
    Ожидает строку вида: API_KEY="ваш_ключ"
    """
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("API_KEY"):
                    # Поддержка формата: API_KEY="ключ" или API_KEY = "ключ"
                    key = line.split("=", 1)[1].strip().strip('"').strip("'")
                    return key
    except FileNotFoundError:
        print(f"Файл {file_path} не найден")
    except Exception as e:
        print(f"Ошибка при чтении API ключа: {e}")
    return None


class AIClassifier:
    """Классификатор текста с улучшенной логикой для валидации документов"""

    def __init__(self, api_key: str = None):
        """Инициализация классификатора"""
        self.api_key = api_key
        self.model = "mistralai/devstral-small:free" #gpt-4o-mini,mistralai/devstral-small:free,moonshotai/kimi-dev-72b:free


        self.valid_classes = [
            'удк', 'автор', 'заголовок', 'сведения_об_авторе',
            'аннотация', 'ключевые_слова',
            'заголовок_английский', 'автор_английский', 'место_работы_английский',
            'аннотация_английская', 'ключевые_слова_английские', 'основной_текст'
        ]

        # Состояние классификации для контекстной логики
        self.classification_state = {
            'title_ru_assigned': False,
            'title_en_assigned': False,
            'abstract_ru_assigned': False,
            'abstract_en_assigned': False,
            'authors_ru_assigned': False,
            'authors_en_assigned': False,
            'keywords_ru_assigned': False,
            'keywords_en_assigned': False,
            'processed_paragraphs': [],
            'current_language_context': 'ru'
        }

    def reset_state(self):
        """Сброс состояния для новой статьи"""
        self.classification_state = {
            'title_ru_assigned': False,
            'title_en_assigned': False,
            'abstract_ru_assigned': False,
            'abstract_en_assigned': False,
            'authors_ru_assigned': False,
            'authors_en_assigned': False,
            'keywords_ru_assigned': False,
            'keywords_en_assigned': False,
            'processed_paragraphs': [],
            'current_language_context': 'ru'
        }

    def classify_paragraph(self, text: str, paragraph_index: int = 0,
                          formatting_info: Dict = None) -> str:
        """
        Улучшенная классификация абзаца с учетом контекста и форматирования

        Args:
            text: Текст абзаца
            paragraph_index: Порядковый номер абзаца в документе
            formatting_info: Информация о форматировании (шрифт, размер, стиль)
        """
        text_clean = text.strip()
        if not text_clean:
            return "основной_текст"

        # Сохраняем информацию об обработанном абзаце
        self.classification_state['processed_paragraphs'].append({
            'index': paragraph_index,
            'text': text_clean[:100],
            'length': len(text_clean)
        })

        # Определяем язык текста
        english_ratio = self._calculate_english_ratio(text_clean)
        is_predominantly_english = english_ratio > 0.7

        # Обновляем языковой контекст
        if is_predominantly_english:
            self.classification_state['current_language_context'] = 'en'
        else:
            self.classification_state['current_language_context'] = 'ru'

        # Применяем правила классификации с учетом контекста
        result = self._classify_with_context(text_clean, paragraph_index,
                                           is_predominantly_english, formatting_info)

        # Обновляем состояние после классификации
        self._update_state_after_classification(result)

        return result

    def _classify_with_context(self, text: str, paragraph_index: int,
                               is_predominantly_english: bool,
                               formatting_info: Dict = None) -> str:
        """Улучшенная классификация с учетом контекста документа"""

        text_lower = text.lower()

        # 1. УДК - всегда в начале документа
        if re.search(r"^удк\s", text, re.IGNORECASE) or paragraph_index <= 2:
            if re.search(r"^удк\s", text, re.IGNORECASE):
                return "удк"

        # # 2. Контекстная классификация авторов
        # author_result = self._classify_authors_with_context(text, is_predominantly_english, paragraph_index)
        # if author_result:
        #     return author_result

        # # 3. ПРИОРИТЕТ: Заголовки должны классифицироваться раньше аннотаций
        # title_result = self._classify_titles_with_context(text, is_predominantly_english, paragraph_index)
        # if title_result:
        #     return title_result

        # 4. Место работы/университеты (для английского текста)
        if is_predominantly_english and self._looks_like_workplace(text):
            return "место_работы_английский"

        # 5. Сведения об авторе (должны идти после авторов)
        if self._is_author_info_context(text, paragraph_index):
            return "сведения_об_авторе"

        # 6. Ключевые слова и аннотации с ключевыми словами (С ОГРАНИЧЕНИЯМИ)
        # keywords_result = self._classify_keywords_and_abstracts(text, text_lower, is_predominantly_english)
        # if keywords_result:
        #     return keywords_result

        # # 7. Аннотации по контексту и длине (ПОСЛЕ проверки заголовков и мест работы)
        # abstract_result = self._classify_abstracts_with_context(text, is_predominantly_english)
        # if abstract_result:
        #     return abstract_result

        # 8. Если правила не дали результата, используем ИИ
        if self.api_key:
            ai_result = self._classify_with_ai(text, is_predominantly_english)
            if ai_result in self.valid_classes:
                return ai_result

        # 9. Резервная классификация
        return self._fallback_classification(text, is_predominantly_english)

    def _classify_authors_with_context(self, text: str, is_english: bool, paragraph_index: int) -> Optional[str]:
        """Улучшенная классификация авторов с учетом контекста"""

        # Авторы обычно идут в начале документа (после УДК, до или после заголовка)
        if paragraph_index > 10:  # Авторы вряд ли будут так далеко
            return None

        if self._looks_like_author(text, is_english):
            if is_english:
                if not self.classification_state['authors_en_assigned']:
                    return "автор_английский"
            else:
                if not self.classification_state['authors_ru_assigned']:
                    return "автор"

        return None

    def _classify_titles_with_context(self, text: str, is_english: bool, paragraph_index: int) -> Optional[str]:
        """Классификация заголовков с учетом контекста и заглавных букв"""

        # Заголовки обычно в начале документа
        if paragraph_index > 8:
            return None

        # Проверка на заголовок с повышенным приоритетом для заглавных букв
        is_title = self._looks_like_title(text)
        is_uppercase_title = self._is_all_uppercase_title(text)

        if is_title or is_uppercase_title:
            if is_english:
                if not self.classification_state['title_en_assigned']:
                    return "заголовок_английский"
            else:
                if not self.classification_state['title_ru_assigned']:
                    return "заголовок"

        return None

    def _is_author_info_context(self, text: str, paragraph_index: int) -> bool:
        """Проверка на сведения об авторе с учетом контекста"""

        # Сведения об авторе обычно идут после списка авторов
        authors_found = any(
            para['text'] for para in self.classification_state['processed_paragraphs'][-3:]
            if self._looks_like_author(para['text'], False) or self._looks_like_author(para['text'], True)
        )

        return (self._looks_like_author_info(text) and
                (authors_found or paragraph_index <= 8))

    def _classify_keywords_and_abstracts(self, text: str, text_lower: str, is_english: bool) -> Optional[str]:
        """Классификация ключевых слов и аннотаций по ключевым словам с ограничениями"""

        # ОГРАНИЧЕНИЯ НА КЛЮЧЕВЫЕ СЛОВА
        # Ключевые слова с явными маркерами (русские)
        if any(keyword in text_lower for keyword in ['Ключевые слова', 'ключевые слова']):
            if not self.classification_state['keywords_ru_assigned']:
                return "ключевые_слова"
            else:
                # Если русские ключевые слова уже назначены, считаем основным текстом
                return "основной_текст"

        # Ключевые слова с явными маркерами (английские)
        if any(keyword in text_lower for keyword in ['Keywords', 'key words']):
            if not self.classification_state['keywords_en_assigned']:
                return "ключевые_слова_английские"
            else:
                # Если английские ключевые слова уже назначены, считаем основным текстом
                return "основной_текст"

        # Аннотации с явными маркерами
        if 'аннотация' or 'статье' in text_lower and not self.classification_state['abstract_ru_assigned']:
            return "аннотация"
        if 'abstract' or 'article' in text_lower and not self.classification_state['abstract_en_assigned']:
            return "аннотация_английская"

        return None

    def _classify_abstracts_with_context(self, text: str, is_english: bool) -> Optional[str]:
        """Классификация аннотаций по контексту и характеристикам"""

        # Критерии для аннотации
        text_length = len(text)
        has_abstract_characteristics = (
            100 <= text_length <= 800 and  # Типичная длина аннотации
            not self._has_structure_words(text) and
            not self._looks_like_author_info(text) and
            not self._has_technical_formulas(text) and
            self._has_abstract_style(text)
        )

        if has_abstract_characteristics:
            if is_english and not self.classification_state['abstract_en_assigned']:
                return "аннотация_английская"
            elif not is_english and not self.classification_state['abstract_ru_assigned']:
                return "аннотация"

        return None

    def _has_abstract_style(self, text: str) -> bool:
        """Проверка стиля аннотации"""
        # Аннотации обычно содержат слова описания исследования
        abstract_indicators = [
            'рассматривается', 'представлен', 'описан', 'исследуется', 'изучается',
            'анализируется', 'предложен', 'разработан', 'получен', 'показано',
            'presents', 'describes', 'analyzes', 'studies', 'investigates',
            'proposes', 'develops', 'demonstrates', 'shows', 'examines',
            'research', 'study', 'analysis', 'investigation', 'method',
            'approach', 'results', 'conclusion', 'findings'
        ]

        text_lower = text.lower()
        return any(indicator in text_lower for indicator in abstract_indicators)

    def _has_technical_formulas(self, text: str) -> bool:
        """Проверка на наличие технических формул"""
        # Простая проверка на математические выражения
        formula_patterns = [
            r'\b[a-zA-Z]\s*[=<>]\s*\d',  # x = 5
            r'\d+\s*[+\-*/]\s*\d+',      # 2 + 3
            r'[а-яА-Я]\s*[=<>]\s*\d'     # х = 5
        ]

        return any(re.search(pattern, text) for pattern in formula_patterns)

    def _update_state_after_classification(self, classification: str):
        """Обновление состояния после классификации"""
        state_map = {
            'заголовок': 'title_ru_assigned',
            'заголовок_английский': 'title_en_assigned',
            'аннотация': 'abstract_ru_assigned',
            'аннотация_английская': 'abstract_en_assigned',
            'автор': 'authors_ru_assigned',
            'автор_английский': 'authors_en_assigned',
            'ключевые_слова': 'keywords_ru_assigned',
            'ключевые_слова_английские': 'keywords_en_assigned'
        }

        if classification in state_map:
            self.classification_state[state_map[classification]] = True

    def _calculate_english_ratio(self, text: str) -> float:
        """Вычисление доли английского текста"""
        english_words = len(re.findall(r'[a-zA-Z]+', text))
        total_words = len(re.findall(r'[а-яёa-zA-Z]+', text, re.IGNORECASE))
        return english_words / max(1, total_words)

    def _looks_like_author(self, text: str, is_english: bool = False) -> bool:
        """Улучшенная проверка на автора"""
        # Проверка на наличие инициалов
        if is_english:
            initials_pattern = r'[A-Z]\.[A-Z]\.'
        else:
            initials_pattern = r'[А-ЯЁ]\.[А-ЯЁ]\.'

        has_initials = bool(re.search(initials_pattern, text))
        if not has_initials:
            return False

        words = text.split()
        if len(words) > 20:  # Слишком длинный для списка авторов
            return False

        # ИСКЛЮЧЕНИЯ
        exclusion_words = [
            'университет', 'институт', 'кафедра', 'доктор', 'профессор', 'доцент',
            'university', 'institute', 'department', 'doctor', 'professor',
            'аннотация', 'abstract', 'ключевые', 'keywords', 'факультет',
            'к.т.н', 'д.т.н', 'заведующий', 'область', 'город', 'улица', 'email', '@'
        ]

        text_lower = text.lower()
        if any(word in text_lower for word in exclusion_words):
            return False

        # Проверка структуры
        author_pattern = r'[А-ЯЁA-Z][а-яёa-z]+\s+[А-ЯЁA-Z]\.[А-ЯЁA-Z]\.'
        authors_found = re.findall(author_pattern, text)
        return len(authors_found) >= 1

    def _looks_like_author_info(self, text: str) -> bool:
        """Проверка на сведения об авторе"""
        full_name_pattern = r'[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+'
        has_full_name = bool(re.search(full_name_pattern, text))

        info_keywords = [
            'кафедра', 'университет', 'институт', 'академия',
            'доктор', 'профессор', 'доцент', 'аспирант',
            'факультет', 'отделение', 'лаборатория',
            'к.т.н', 'д.т.н', 'заведующий', 'область', 'город', 'улица', 'email', '@'
        ]

        text_lower = text.lower()
        has_info_keywords = any(keyword in text_lower for keyword in info_keywords)
        has_postal_code = bool(re.search(r'\d{6}', text))
        has_professional_abbr = bool(re.search(r'[кд]\.т\.н', text_lower))

        return has_full_name or has_info_keywords or has_postal_code or has_professional_abbr

    def _looks_like_title(self, text: str) -> bool:
        """Более строгая проверка на заголовок"""
        words = text.split()

        # Базовые критерии
        if not (3 <= len(words) <= 20):
            return False

        # Исключаем инициалы и email
        if '@' in text or re.search(r'[А-ЯЁA-Z]\.[А-ЯЁA-Z]\.', text):
            return False

        # Исключаем университеты и организации
        if self._looks_like_workplace(text):
            return False

        # Исключаем адреса и города
        if self._has_address_pattern(text):
            return False

        # Исключаем профессиональную информацию
        if self._has_professional_keywords(text):
            return False

        # Заголовки часто содержат предлоги и артикли
        title_indicators = ['the', 'of', 'in', 'on', 'for', 'with', 'by', 'and', 'or']
        text_lower = text.lower()
        has_title_words = any(word in text_lower for word in title_indicators)

        location_words = ['russia', 'moscow', 'university', 'institute', 'academy', 'center', 'centre']
        has_location_words = any(word in text_lower for word in location_words)

        # Заголовок должен содержать слова-индикаторы, но не географические названия
        return has_title_words and not has_location_words

    def _is_all_uppercase_title(self, text: str) -> bool:
        """Проверка на заголовок, написанный заглавными буквами"""
        # Убираем знаки препинания и пробелы для анализа
        letters_only = re.sub(r'[^\w\s]', '', text)

        # Подсчитываем буквы
        uppercase_letters = len(re.findall(r'[А-ЯЁA-Z]', letters_only))
        lowercase_letters = len(re.findall(r'[а-яёa-z]', letters_only))
        total_letters = uppercase_letters + lowercase_letters

        # Если большинство букв заглавные (>80%), это вероятно заголовок
        if total_letters > 0:
            uppercase_ratio = uppercase_letters / total_letters
            return uppercase_ratio > 0.8

        return False

    def _has_structure_words(self, text: str) -> bool:
        """Проверка на наличие структурных слов"""
        structure_words = [
            'введение', 'заключение', 'выводы', 'методы', 'результаты',
            'обсуждение', 'литература', 'список', 'библиография',
            'introduction', 'conclusion', 'methods', 'results', 'discussion'
        ]
        text_lower = text.lower()
        return any(word in text_lower for word in structure_words)

    def _has_address_pattern(self, text: str) -> bool:
        """Проверка на адресную информацию"""
        if re.search(r'\d{6}', text):
            return True
        address_keywords = [r'область', r'город', r'улица', r'дом', r'ул\.', r'г\.', r'д\.']
        text_lower = text.lower()
        return any(keyword in text_lower for keyword in address_keywords)

    def _has_professional_keywords(self, text: str) -> bool:
        """Проверка на профессиональные ключевые слова"""
        prof_keywords = [
            'к.т.н', 'д.т.н', 'доктор', 'кандидат', 'профессор', 'доцент',
            'заведующий', 'кафедра', 'университет', 'институт'
        ]
        text_lower = text.lower()
        return any(keyword in text_lower for keyword in prof_keywords)

    def _classify_with_ai(self, text: str, is_predominantly_english: bool = False,
                         max_retries: int = 3) -> str:
        """Классификация с помощью ИИ через OpenRouter API"""

        if not self.api_key:
            return self._fallback_classification(text, is_predominantly_english)

        # Определяем релевантные классы с учетом ограничений
        if is_predominantly_english:
            relevant_classes = [
                'заголовок_английский', 'автор_английский', 'место_работы_английский',
                'аннотация_английская'
            ]
            # Добавляем ключевые слова только если они еще не назначены
            if not self.classification_state['keywords_en_assigned']:
                relevant_classes.append('ключевые_слова_английские')
        else:
            relevant_classes = [
                'автор', 'заголовок', 'сведения_об_авторе',
                'аннотация', 'основной_текст'
            ]
            # Добавляем ключевые слова только если они еще не назначены
            if not self.classification_state['keywords_ru_assigned']:
                relevant_classes.append('ключевые_слова')

        # Создаем контекстный промпт
        context_info = self._build_context_for_ai()

        prompt = f"""Определи тип элемента научной статьи. Ответь ТОЛЬКО одним словом из списка:
{', '.join(relevant_classes)}

Контекст документа:
{context_info}

Правила классификации:
1. автор/автор_английский:
   - Строка с фамилиями и инициалами авторов
   - Пример: "Иванов А.А., Петров Б.В."
   - Может содержать email (но не всегда)

2. заголовок/заголовок_английский:
   - Название статьи (обычно 5-20 слов)
   - Часто выделен жирным или заглавными буквами
   - Содержит ключевые термины исследования
   - В русской статье английский заголовок идет после русского

3. сведения_об_авторе/место_работы_английский:
   - Информация об аффилиации, должностях, ученых степенях
   - Содержит названия организаций, городов
   - Может включать контактную информацию
   - Пример: "МГУ им. Ломоносова, Москва, Россия"

4. аннотация/аннотация_английская:
   - Краткое описание исследования (100-500 символов)
   - Начинается с "В статье..." или "The article..."
   - Содержит цели, методы и основные результаты

5. ключевые_слова/ключевые_слова_английские:
   - Начинаются со слов "Ключевые слова:" или "Keywords:"
   - Содержат 3-10 терминов через запятую

6. основной_текст:
   - Содержит научное описание исследования
   - Включает формулы, ссылки на литературу
   - Может содержать подзаголовки (введение, методы и т.д.)
   - Часто использует научную лексику

7. удостоверяющая_информация:
   - УДК, DOI, дата поступления
   - Пример: "УДК 66.02:519.771.3"
   

Текст: "{text[:500]}"

Тип:"""

        for attempt in range(max_retries):
            try:
                response = requests.post(
                    url="https://openrouter.ai/api/v1/chat/completions",
                    headers={
                        "Authorization": f"Bearer {self.api_key}",
                        "Content-Type": "application/json",
                        "X-Title": "Scientific Text Classifier"
                    },
                    data=json.dumps({
                        "model": self.model,
                        "messages": [{"role": "user", "content": prompt}],
                        "temperature": 0.1,
                        "max_tokens": 10,
                        "top_p": 0.3
                    }),
                    timeout=15
                )

                if response.status_code == 200:
                    result = response.json()['choices'][0]['message']['content'].strip().lower()

                    # Поиск подходящего класса с проверкой ограничений
                    for valid_class in relevant_classes:
                        if valid_class.lower() in result or result in valid_class.lower():
                            # Дополнительная проверка для ключевых слов
                            if valid_class == 'ключевые_слова' and self.classification_state['keywords_ru_assigned']:
                                continue
                            if valid_class == 'ключевые_слова_английские' and self.classification_state['keywords_en_assigned']:
                                continue
                            return valid_class

                    return self._fallback_classification(text, is_predominantly_english)

                else:
                    if response.status_code == 429:
                        sleep(2 ** attempt)
                    elif attempt < max_retries - 1:
                        sleep(1)

            except Exception as e:
                if attempt < max_retries - 1:
                    sleep(2)

        return self._fallback_classification(text, is_predominantly_english)

    def _build_context_for_ai(self) -> str:
        """Создание контекстной информации для ИИ"""
        context_parts = []

        if self.classification_state['title_ru_assigned']:
            context_parts.append("русский заголовок уже найден")
        if self.classification_state['title_en_assigned']:
            context_parts.append("английский заголовок уже найден")
        if self.classification_state['abstract_ru_assigned']:
            context_parts.append("русская аннотация уже найдена")
        if self.classification_state['abstract_en_assigned']:
            context_parts.append("английская аннотация уже найдена")
        if self.classification_state['keywords_ru_assigned']:
            context_parts.append("русские ключевые слова уже найдены")
        if self.classification_state['keywords_en_assigned']:
            context_parts.append("английские ключевые слова уже найдены")

        context_info = "; ".join(context_parts) if context_parts else "начало документа"
        return context_info

    def _fallback_classification(self, text: str, is_predominantly_english: bool = False) -> str:
        """Резервная классификация без ИИ с учетом ограничений"""
        if is_predominantly_english:
            return self._classify_english_text(text)
        else:
            return self._classify_russian_text(text)

    def _classify_english_text(self, text: str) -> str:
        """Исправленная классификация английского текста"""

        # 1. Сначала проверяем на заголовок
        if (self._looks_like_title(text) and
                not self.classification_state['title_en_assigned'] and
                not self._looks_like_workplace(text)):
            return "заголовок_английский"
        # 2. Проверяем на место работы/университет
        elif self._looks_like_workplace(text):
            return "место_работы_английский"

        # 3. Проверяем на аннотацию (с дополнительными критериями)
        elif (100 <= len(text) <= 600 and
              not self.classification_state['abstract_en_assigned'] and
              self._has_abstract_style(text) and
              not self._looks_like_title(text) and
              not self._looks_like_workplace(text)):
            return "аннотация_английская"

        # 4. Ключевые слова
        elif (',' in text and len(text) <= 100 and
              not self.classification_state['keywords_en_assigned']):
            return "ключевые_слова_английские"

        else:
            return "основной_текст"

    def _classify_russian_text(self, text: str) -> str:
        """Классификация русского текста с ограничениями"""
        if (self._looks_like_title(text) and
            not self._looks_like_author_info(text) and
            not self.classification_state['title_ru_assigned']):
            return "заголовок"
        elif self._has_address_pattern(text) or '@' in text:
            return "сведения_об_авторе"
        elif (100 <= len(text) <= 600 and
              self._has_abstract_style(text) and
              not self._has_structure_words(text) and
              not self.classification_state['abstract_ru_assigned']):
            return "аннотация"
        elif (len(text) <= 100 and ',' in text and
              not self._looks_like_author_info(text) and
              not self.classification_state['keywords_ru_assigned']):
            return "ключевые_слова"
        else:
            return "основной_текст"

    def _looks_like_workplace(self, text: str) -> bool:
        """Улучшенная проверка на место работы"""
        workplace_keywords = [
            'university', 'institute', 'academy', 'center', 'centre',
            'department', 'faculty', 'school', 'college', 'laboratory',
            'company', 'corporation', 'ltd', 'inc', 'llc'
        ]

        # Географические названия
        location_keywords = [
            'russia', 'moscow', 'petersburg', 'novomoskovsk', 'usa', 'uk',
            'germany', 'france', 'china', 'japan', 'street', 'avenue', 'road'
        ]

        text_lower = text.lower()

        # Проверяем наличие ключевых слов организаций
        has_workplace_keywords = any(keyword in text_lower for keyword in workplace_keywords)

        # Проверяем наличие географических названий
        has_location_keywords = any(keyword in text_lower for keyword in location_keywords)

        # Проверяем наличие имени собственного (заглавные буквы)
        has_proper_nouns = bool(re.search(r'[A-Z][a-z]+', text))

        return has_workplace_keywords or (has_location_keywords and has_proper_nouns)


# Совместимость с существующим кодом
def reset_flags(self):
    """Метод для совместимости с существующим кодом"""
    self.reset_state()

# Добавляем метод для совместимости
AIClassifier.reset_flags = reset_flags