# -*- coding: utf-8 -*-
# import numpy as np
# import pandas as pd
import re
from docx import Document
import nltk
from nltk.corpus import stopwords
from collections import Counter
import string

# стоп-слова для русского языка
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    nltk.download('punkt')
try:
    nltk.data.find('corpora/stopwords')
except LookupError:
    nltk.download('stopwords')
russian_stopwords = set(stopwords.words('russian'))

# специфические для текста стоп-слова (знаки препинания, лишние символы)
# и возможные артефакты разметки
custom_stopwords = set(['это', 'что', 'как', 'он', 'она', 'они', 'я', 'ты', 'мы', 'вы',
                        'его', 'ее', 'их', 'мой', 'твой', 'наш', 'ваш', 'весь', 'эти',
                        'который', 'такой', 'себя', '—', '...', '..', '.', ',', '!', '?',
                        ';', ':', '(', ')', '[', ']', '№', '-', 'бы', 'же', 'ли', 'вот',
                        'уже', 'еще', 'когда', 'только', 'было', 'был', 'была', 'были',
                        'очень', 'совсем', 'можно', 'надо', 'нужно', 'чтобы', 'также',
                        'например', 'именно'])

all_stopwords = russian_stopwords.union(custom_stopwords)

def clean_text(text):
    text = text.lower()
    # Убираем цифры
    text = re.sub(r'\d+', '', text)
    # Убираем пунктуацию (оставляем только буквы и пробелы)
    text = text.translate(str.maketrans('', '', string.punctuation))
    # Убираем лишние пробелы
    text = re.sub(r'\s+', ' ', text).strip()
    return text

# Функция для токенизации и удаления стоп-слов
def tokenize_and_filter(text):
    tokens = text.split()
    filtered_tokens = [word for word in tokens if word not in all_stopwords and len(word) > 1]
    return filtered_tokens

def parse_docx_sentiment(filepath):
    """
    читает docx файл и возвращает список предложений с тональностью.
    игнорирует строки без явной разметки в конце.
    """
    try:
        doc = Document(filepath)
    except Exception as e:
        print(f"Ошибка при открытии файла: {e}")
        return []

    sentences_data = []
    # паттерн для поиска тональности в конце строки: [1], [2], [3] или [4]
    pattern = r'\[(\d)\]\.?\s*$'

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # Ищем маркер тональности в конце строки
        match = re.search(pattern, text)
        if match:
            sentiment = int(match.group(1))
            # Убираем сам маркер из текста предложения
            clean_sentence = re.sub(r'\s*\[\d\]\.?\s*$', '', text).strip()
            if clean_sentence: # Добавляем только непустые предложения
                sentences_data.append({
                    'text': clean_sentence,
                    'sentiment': sentiment
                })
    print(f"Найдено размеченных предложений: {len(sentences_data)}")
    return sentences_data

# 2. Создаем словари частотности для каждой тональности
def build_sentiment_dictionaries(sentences_data):
    """
    создает 4 словаря (Counter) для тональностей 1,2,3,4.
    считает частоту слов в каждой категории после очистки.
    """
    # Инициализируем пустые счетчики для каждой тональности
    sentiment_dicts = {
        1: Counter(),
        2: Counter(),
        3: Counter(),
        4: Counter()
    }

    # Для подсчета общего количества предложений в каждом классе (опционально)
    class_counts = {1:0, 2:0, 3:0, 4:0}

    for item in sentences_data:
        sentiment = item['sentiment']
        raw_text = item['text']

        # Очищаем текст и токенизируем
        cleaned = clean_text(raw_text)
        tokens = tokenize_and_filter(cleaned)

        # Обновляем счетчик для данного класса тональности
        sentiment_dicts[sentiment].update(tokens)
        class_counts[sentiment] += 1

    print("\nСтатистика по классам:")
    t = ["Отрицательная", "Нейтральная", "Положительная", "Неоднозначная"]
    i = 0
    for sent, count in class_counts.items():
        print(f"{t[i]} тональность: {count} предложений")
        i += 1

    return sentiment_dicts

# Главная функция, которая собирает всё вместе
def main(docx_path):
    print("Запуск анализа тональности...")
    # Шаг 1: Парсим
    data = parse_docx_sentiment(docx_path)

    if not data:
        print("Нет данных для обработки. Завершение.")
        return

    # Шаг 2 и 3: Строим словари
    sentiment_dicts = build_sentiment_dictionaries(data)
    return sentiment_dicts

if __name__ == "__main__":
    file_path = "разметка морфий-1.docx"
    sentiment_dicts = main(file_path)
