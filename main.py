# -*- coding: utf-8 -*-
import re
from docx import Document
import nltk
from nltk.corpus import stopwords
from collections import Counter
import string
import pymorphy3

try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    nltk.download('punkt')
try:
    nltk.data.find('corpora/stopwords')
except LookupError:
    nltk.download('stopwords')

russian_stopwords = set(stopwords.words('russian'))

custom_stopwords = set(['это', 'что', 'как', 'он', 'она', 'они', 'я', 'ты', 'мы', 'вы',
                        'его', 'ее', 'их', 'мой', 'твой', 'наш', 'ваш', 'весь', 'эти',
                        'который', 'такой', 'себя', '—', '...', '..', '.', ',', '!', '?',
                        ';', ':', '(', ')', '[', ']', '№', '-', 'бы', 'же', 'ли', 'вот',
                        'уже', 'еще', 'когда', 'только', 'было', 'был', 'была', 'были',
                        'очень', 'совсем', 'можно', 'надо', 'нужно', 'чтобы', 'также',
                        'например', 'именно'])

all_stopwords = russian_stopwords.union(custom_stopwords)

morph = pymorphy3.MorphAnalyzer()


def clean_text(text):
    text = text.lower()
    text = re.sub(r'\d+', '', text)
    text = text.translate(str.maketrans('', '', string.punctuation))
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def tokenize_and_filter(text):
    tokens = text.split()
    filtered_tokens = [word for word in tokens if word not in all_stopwords and len(word) > 1]
    return filtered_tokens


def get_verb_mood(word):
    try:
        parsed = morph.parse(word)[0]
        if parsed.tag.POS == 'VERB':
            mood = parsed.tag.mood
            if mood == 'indc':
                return 'изъявительное'
            elif mood == 'impr':
                return 'повелительное'
            else:
                return 'изъявительное'
        return None
    except:
        return None


def extract_verbs_with_mood_from_text(text):
    cleaned = clean_text(text)
    words = cleaned.split()
    verbs = []
    for word in words:
        mood = get_verb_mood(word)
        if mood is not None:
            verbs.append(mood)
    return verbs


def parse_docx_all_lines(filepath):
    try:
        doc = Document(filepath)
    except Exception as e:
        print(f"Ошибка при открытии файла: {e}")
        return []

    lines_data = []
    pattern = r'\[(\d)\]\.?\s*$'

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        match = re.search(pattern, text)
        if match:
            sentiment = int(match.group(1))
            clean_sentence = re.sub(r'\s*\[\d\]\.?\s*$', '', text).strip()
            lines_data.append({
                'text': clean_sentence,
                'sentiment': sentiment,
                'is_sentence': True
            })
        else:
            lines_data.append({
                'text': text,
                'sentiment': None,
                'is_sentence': False
            })

    print(
        f"Найдено строк: {len(lines_data)} (из них предложений с разметкой: {sum(1 for x in lines_data if x['is_sentence'])})")
    return lines_data


def split_into_chapters_from_lines(lines_data):
    chapters = {}
    current_chapter = 0
    chapter_sentences = []

    for item in lines_data:
        text = item['text']

        clean_text_for_check = text.strip('*').strip()
        chapter_match = re.match(r'^Глава\s+(\d+)', clean_text_for_check, re.IGNORECASE)

        if chapter_match and not item['is_sentence']:
            if current_chapter > 0 and chapter_sentences:
                chapters[current_chapter] = chapter_sentences.copy()
                print(f"Глава {current_chapter}: {len(chapter_sentences)} предложений")
            current_chapter = int(chapter_match.group(1))
            chapter_sentences = []
            print(f"Найдена глава {current_chapter}")
        elif item['is_sentence'] and current_chapter > 0:
            chapter_sentences.append({
                'text': item['text'],
                'sentiment': item['sentiment']
            })

    if current_chapter > 0 and chapter_sentences:
        chapters[current_chapter] = chapter_sentences
        print(f"Глава {current_chapter}: {len(chapter_sentences)} предложений")

    return chapters


def analyze_chapter_sentences(chapter_sentences):
    counts = {1: 0, 2: 0, 3: 0, 4: 0}
    total = len(chapter_sentences)

    for item in chapter_sentences:
        sentiment = item['sentiment']
        counts[sentiment] += 1

    shares = {}
    for sentiment in [1, 2, 3, 4]:
        shares[sentiment] = (counts[sentiment] / total * 100) if total > 0 else 0

    return counts, shares, total


def analyze_chapter_verbs_by_mood(chapter_sentences):
    mood_counts = {
        'изъявительное': 0,
        'повелительное': 0,
        'сослагательное': 0
    }
    total_verbs = 0

    for item in chapter_sentences:
        text = item['text']
        verbs = extract_verbs_with_mood_from_text(text)
        for mood in verbs:
            mood_counts[mood] += 1
            total_verbs += 1

    shares = {}
    for mood in ['изъявительное', 'повелительное', 'сослагательное']:
        shares[mood] = (mood_counts[mood] / total_verbs * 100) if total_verbs > 0 else 0

    return mood_counts, shares, total_verbs


def print_chapter_results(chapters):
    print("\n")
    print("АНАЛИЗ ПО ГЛАВАМ")

    sentiment_names = {
        1: "Отрицательная",
        2: "Нейтральная",
        3: "Положительная",
        4: "Неоднозначная"
    }

    mood_names = {
        'изъявительное': "Изъявительное",
        'повелительное': "Повелительное",
        'сослагательное': "Сослагательное"
    }

    for chapter_num in sorted(chapters.keys()):
        chapter_sentences = chapters[chapter_num]

        sent_counts, sent_shares, sent_total = analyze_chapter_sentences(chapter_sentences)

        verb_counts, verb_shares, verb_total = analyze_chapter_verbs_by_mood(chapter_sentences)

        print(f"\n")
        print(f"ГЛАВА {chapter_num}")
        print(f"Всего предложений в главе: {sent_total}")
        print(f"Всего глаголов в главе: {verb_total}")

        print(f"\n--- РАСПРЕДЕЛЕНИЕ ПРЕДЛОЖЕНИЙ ПО ТОНАЛЬНОСТИ ---")
        print(f"{'Тональность':<20} {'Кол-во':<10} {'Доля':<10}")
        print("-" * 45)
        for sentiment in [1, 2, 3, 4]:
            count = sent_counts[sentiment]
            share = sent_shares[sentiment]
            print(f"{sentiment_names[sentiment]:<20} {count:<10} {share:.2f}%")
        print("-" * 45)
        print(f"{'ИТОГО':<20} {sent_total:<10} 100.00%")

        print(f"\n--- РАСПРЕДЕЛЕНИЕ ГЛАГОЛОВ ПО НАКЛОНЕНИЮ ---")
        print(f"{'Наклонение':<20} {'Кол-во':<10} {'Доля':<10}")
        print("-" * 45)
        for mood in ['изъявительное', 'повелительное', 'сослагательное']:
            count = verb_counts[mood]
            share = verb_shares[mood]
            print(f"{mood_names[mood]:<20} {count:<10} {share:.2f}%")
        print("-" * 45)
        print(f"{'ИТОГО':<20} {verb_total:<10} 100.00%")


def print_total_results(data):
    sent_counts = {1: 0, 2: 0, 3: 0, 4: 0}
    for item in data:
        sent_counts[item['sentiment']] += 1
    sent_total = len(data)

    verb_counts = {'изъявительное': 0, 'повелительное': 0, 'сослагательное': 0}
    verb_total = 0
    for item in data:
        verbs = extract_verbs_with_mood_from_text(item['text'])
        for mood in verbs:
            verb_counts[mood] += 1
            verb_total += 1

    sentiment_names = {
        1: "Отрицательная",
        2: "Нейтральная",
        3: "Положительная",
        4: "Неоднозначная"
    }

    mood_names = {
        'изъявительное': "Изъявительное",
        'повелительное': "Повелительное",
        'сослагательное': "Сослагательное"
    }

    print("\n")
    print("ОБЩИЕ РЕЗУЛЬТАТЫ ПО ВСЕМУ ТЕКСТУ")

    print(f"\nВсего предложений в тексте: {sent_total}")
    print(f"Всего глаголов в тексте: {verb_total}")

    print(f"\n--- РАСПРЕДЕЛЕНИЕ ПРЕДЛОЖЕНИЙ ПО ТОНАЛЬНОСТИ ---")
    print(f"{'Тональность':<20} {'Кол-во':<10} {'Доля':<10}")
    print("-" * 45)
    for sentiment in [1, 2, 3, 4]:
        count = sent_counts[sentiment]
        share = (count / sent_total * 100) if sent_total > 0 else 0
        print(f"{sentiment_names[sentiment]:<20} {count:<10} {share:.2f}%")
    print("-" * 45)
    print(f"{'ИТОГО':<20} {sent_total:<10} 100.00%")

    print(f"\n--- РАСПРЕДЕЛЕНИЕ ГЛАГОЛОВ ПО НАКЛОНЕНИЮ ---")
    print(f"{'Наклонение':<20} {'Кол-во':<10} {'Доля':<10}")
    print("-" * 45)
    for mood in ['изъявительное', 'повелительное', 'сослагательное']:
        count = verb_counts[mood]
        share = (count / verb_total * 100) if verb_total > 0 else 0
        print(f"{mood_names[mood]:<20} {count:<10} {share:.2f}%")
    print("-" * 45)
    print(f"{'ИТОГО':<20} {verb_total:<10} 100.00%")


def main(docx_path):
    print("Запуск анализа...")

    lines = parse_docx_all_lines(docx_path)
    if not lines:
        print("Нет данных для обработки. Завершение.")
        return

    data = [x for x in lines if x['is_sentence']]
    if not data:
        print("Нет размеченных предложений. Завершение.")
        return

    chapters = split_into_chapters_from_lines(lines)
    print(f"\nНайдено глав: {len(chapters)}")

    print_total_results(data)
    print_chapter_results(chapters)

    return chapters


if __name__ == "__main__":
    file_path = "разметка морфий-1.docx"
    chapters = main(file_path)