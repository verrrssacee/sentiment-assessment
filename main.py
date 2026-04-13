# -*- coding: utf-8 -*-
import re
from docx import Document
import nltk
from nltk.corpus import stopwords
from collections import Counter
import string
import pymorphy3

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import accuracy_score, classification_report, confusion_matrix
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

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

        print(f"\nРАСПРЕДЕЛЕНИЕ ПРЕДЛОЖЕНИЙ ПО ТОНАЛЬНОСТИ")
        print(f"{'Тональность':<20} {'Кол-во':<10} {'Доля':<10}")
        print("-" * 45)
        for sentiment in [1, 2, 3, 4]:
            count = sent_counts[sentiment]
            share = sent_shares[sentiment]
            print(f"{sentiment_names[sentiment]:<20} {count:<10} {share:.2f}%")
        print("-" * 45)
        print(f"{'ИТОГО':<20} {sent_total:<10} 100.00%")

        print(f"\nРАСПРЕДЕЛЕНИЕ ГЛАГОЛОВ ПО НАКЛОНЕНИЮ")
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

    print(f"\nРАСПРЕДЕЛЕНИЕ ПРЕДЛОЖЕНИЙ ПО ТОНАЛЬНОСТИ")
    print(f"{'Тональность':<20} {'Кол-во':<10} {'Доля':<10}")
    print("-" * 45)
    for sentiment in [1, 2, 3, 4]:
        count = sent_counts[sentiment]
        share = (count / sent_total * 100) if sent_total > 0 else 0
        print(f"{sentiment_names[sentiment]:<20} {count:<10} {share:.2f}%")
    print("-" * 45)
    print(f"{'ИТОГО':<20} {sent_total:<10} 100.00%")

    print(f"\nРАСПРЕДЕЛЕНИЕ ГЛАГОЛОВ ПО НАКЛОНЕНИЮ")
    print(f"{'Наклонение':<20} {'Кол-во':<10} {'Доля':<10}")
    print("-" * 45)
    for mood in ['изъявительное', 'повелительное', 'сослагательное']:
        count = verb_counts[mood]
        share = (count / verb_total * 100) if verb_total > 0 else 0
        print(f"{mood_names[mood]:<20} {count:<10} {share:.2f}%")
    print("-" * 45)
    print(f"{'ИТОГО':<20} {verb_total:<10} 100.00%")


def prepare_texts_and_labels(data):
    """
    Подготавливает тексты и метки для машинного обучения.
    Возвращает список текстов и список меток.
    """
    texts = [item['text'] for item in data]
    labels = [item['sentiment'] for item in data]
    return texts, labels


def create_tfidf_matrix(texts, max_features=5000, ngram_range=(1, 2)):
    """
    Создаёт TF-IDF матрицу из списка текстов.
    Возвращает матрицу признаков и векторизатор.
    """
    vectorizer = TfidfVectorizer(
        max_features=max_features,
        ngram_range=ngram_range,
        min_df=2,
        max_df=0.9
    )
    X = vectorizer.fit_transform(texts)
    print(f"TF-IDF матрица создана. Размер: {X.shape}")
    return X, vectorizer


def train_logistic_regression(X_train, y_train):
    """Обучает модель логистической регрессии"""
    model = LogisticRegression(
        max_iter=1000,
        C=1.0,
        random_state=42
    )
    model.fit(X_train, y_train)
    print("Логистическая регрессия обучена")
    return model


def train_decision_tree(X_train, y_train):
    """Обучает модель дерева решений"""
    model = DecisionTreeClassifier(
        max_depth=10,
        min_samples_split=5,
        min_samples_leaf=2,
        random_state=42
    )
    model.fit(X_train, y_train)
    print("Дерево решений обучено")
    return model


def evaluate_model(model, X_test, y_test, model_name):
    """
    Оценивает модель и выводит метрики качества.
    Возвращает accuracy.
    """
    y_pred = model.predict(X_test)
    accuracy = accuracy_score(y_test, y_pred)

    print(f"\n{model_name}")
    print(f"Accuracy: {accuracy:.4f}")
    print("\nClassification Report:")
    print(classification_report(y_test, y_pred))

    # Матрица ошибок
    cm = confusion_matrix(y_test, y_pred)
    plt.figure(figsize=(8, 6))
    sns.heatmap(cm, annot=True, fmt='d', cmap='Blues',
                xticklabels=['Отриц.', 'Нейтр.', 'Пол.', 'Неодн.'],
                yticklabels=['Отриц.', 'Нейтр.', 'Пол.', 'Неодн.'])
    plt.title(f'Матрица ошибок - {model_name}')
    plt.xlabel('Предсказанный класс')
    plt.ylabel('Истинный класс')
    plt.tight_layout()
    plt.savefig(f'confusion_matrix_{model_name}.png')
    plt.show()

    return accuracy, y_pred


def compare_models(models_results):
    """
    Сравнивает результаты нескольких моделей.
    models_results: список словарей с ключами 'name', 'accuracy'
    """
    df = pd.DataFrame(models_results)
    print("\nСРАВНИТЕЛЬНЫЙ АНАЛИЗ МОДЕЛЕЙ")
    print(df.to_string(index=False))

    # Визуализация сравнения
    plt.figure(figsize=(8, 5))
    bars = plt.bar(df['name'], df['accuracy'], color=['blue', 'green'])
    plt.xlabel('Модель')
    plt.ylabel('Accuracy')
    plt.title('Сравнение точности моделей')
    plt.ylim(0, 1)
    for bar, acc in zip(bars, df['accuracy']):
        plt.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.01,
                 f'{acc:.4f}', ha='center', va='bottom')
    plt.tight_layout()
    plt.savefig('models_comparison.png')
    plt.show()


def run_ml_experiment(data):
    """
    Запускает эксперимент по машинному обучению:
    - векторизация текстов
    - обучение моделей
    - оценка качества
    - сравнительный анализ
    """
    print("ЭКСПЕРИМЕНТ ПО МАШИННОМУ ОБУЧЕНИЮ")

    # Подготовка данных
    texts, labels = prepare_texts_and_labels(data)
    print(f"Всего примеров: {len(texts)}")

    # TF-IDF векторизация
    X, vectorizer = create_tfidf_matrix(texts)

    # Разделение на обучающую и тестовую выборки (80% / 20%)
    X_train, X_test, y_train, y_test = train_test_split(
        X, labels, test_size=0.2, random_state=42, stratify=labels
    )
    print(f"Обучающая выборка: {X_train.shape[0]} примеров")
    print(f"Тестовая выборка: {X_test.shape[0]} примеров")

    # Обучение моделей
    lr_model = train_logistic_regression(X_train, y_train)
    dt_model = train_decision_tree(X_train, y_train)

    # Оценка моделей
    acc_lr, _ = evaluate_model(lr_model, X_test, y_test, "Логистическая регрессия")
    acc_dt, _ = evaluate_model(dt_model, X_test, y_test, "Дерево решений")

    # Сравнительный анализ
    results = [
        {'name': 'Логистическая регрессия', 'accuracy': acc_lr},
        {'name': 'Дерево решений', 'accuracy': acc_dt}
    ]
    compare_models(results)

    return lr_model, dt_model, vectorizer


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

    lr_model, dt_model, vectorizer = run_ml_experiment(data)

    return chapters, lr_model, dt_model, vectorizer


if __name__ == "__main__":
    file_path = "разметка морфий-1.docx"
    chapters = main(file_path)
