import os
import re
import joblib
import pandas as pd
import numpy as np

import nltk
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer

from docx import Document
import pdfplumber
import win32com.client

from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.pipeline import Pipeline
from sklearn.feature_extraction.text import TfidfVectorizer

from sklearn.linear_model import LogisticRegression
from sklearn.svm import LinearSVC
from sklearn.naive_bayes import MultinomialNB
from sklearn.neighbors import KNeighborsClassifier

from sklearn.metrics import classification_report, accuracy_score



nltk.download('stopwords')
nltk.download('wordnet')

stop_words = set(stopwords.words('english'))
lemmatizer = WordNetLemmatizer()


def clean_text(text):
    text = str(text).lower()
    text = re.sub(r'http\S+|www\S+', ' ', text)
    text = re.sub(r'[^a-z\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text)

    words = text.split()
    words = [w for w in words if w not in stop_words and len(w) > 2]
    words = [lemmatizer.lemmatize(w) for w in words]

    return " ".join(words)


def read_docx(path):
    doc = Document(path)
    return "\n".join([p.text for p in doc.paragraphs])

def read_pdf(path):
    text = ""
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ""
    return text

def read_doc(path, word):
    doc = word.Documents.Open(path)
    text = doc.Content.Text
    doc.Close()
    return text


def classify_resume(text):
    text = text.lower()

    if "peoplesoft" in text:
        return "Peoplesoft"
    elif "workday" in text:
        return "Workday"
    elif "reactjs" in text:
        return "Reactjs"
    elif "react" in text:
        return "React"
    elif "sql" in text:
        return "SQL"
    else:
        return "Internship"


def load_data(path):
    texts, labels = [], []

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    for root, _, files in os.walk(path):
        for file in files:
            file_path = os.path.join(root, file)
            file_lower = file.lower()

            text = ""

            if file_lower.endswith(".pdf"):
                text = read_pdf(file_path)
            elif file_lower.endswith(".docx"):
                text = read_docx(file_path)
            elif file_lower.endswith(".doc"):
                text = read_doc(file_path, word)
            else:
                continue

            if text.strip():
                cleaned = clean_text(text)
                label = classify_resume(cleaned)

                texts.append(cleaned)
                labels.append(label)

    word.Quit()

    df = pd.DataFrame({"text": texts, "label": labels})

    print("\nDataset Shape:", df.shape)
    print("\nCategory Distribution:\n", df["label"].value_counts())

    return df


def train_model(path):

    df = load_data(path)

    # ✅ Merge categories into "Other" (YOUR REQUIREMENT)
    df["label"] = df["label"].replace({
        "React": "Other",
        "Reactjs": "Other",
        "Internship": "Other"
    })

    print("\nUpdated Category Distribution:\n", df["label"].value_counts())

    X = df["text"]
    y = df["label"]

    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, stratify=y, random_state=42
    )

    models = {
        "Logistic Regression": LogisticRegression(max_iter=1000),
        "Naive Bayes": MultinomialNB(),
        "Linear SVM": LinearSVC(),
        "KNN": KNeighborsClassifier()
    }

    param_grids = {
        "Logistic Regression": {
            'tfidf__max_features': [3000, 5000],
            'tfidf__ngram_range': [(1,1), (1,2)],
            'clf__C': [0.1, 1, 10]
        },

        "Naive Bayes": {
            'tfidf__max_features': [3000, 5000],
            'tfidf__ngram_range': [(1,1), (1,2)],
            'clf__alpha': [0.5, 1.0]
        },

        "Linear SVM": {
            'tfidf__max_features': [300, 500],
            'clf__C': [0.001, 0.01],
            'clf__class_weight': [None, 'balanced']
        },

        "KNN": {
            'tfidf__max_features': [5000],
            'clf__n_neighbors': [5, 7]
        }
    }

    results = {}
    best_models = {}

    for name, model in models.items():

        print(f"\n🚀 Training {name}")

        pipeline = Pipeline([
            ("tfidf", TfidfVectorizer()),
            ("clf", model)
        ])

        grid = GridSearchCV(
            pipeline,
            param_grids[name],
            cv=5,
            scoring='accuracy',
            n_jobs=-1
        )

        grid.fit(X_train, y_train)

        best_models[name] = grid.best_estimator_

        preds = grid.predict(X_test)

        test_acc = accuracy_score(y_test, preds)
        cv_score = grid.best_score_

        results[name] = cv_score

        print(f"\n{name} Results:")
        print("Best CV Score:", cv_score)
        print("Test Accuracy:", test_acc)
        print(classification_report(y_test, preds))

   
    results_df = pd.DataFrame({
        "Model": results.keys(),
        "Best CV Score": results.values()
    }).sort_values(by="Best CV Score", ascending=False)

    print("\n📊 Final Model Comparison:\n")
    print(results_df)

    best_model_name = results_df.iloc[0]["Model"]
    best_model = best_models[best_model_name]

    print(f"\n🏆 Best Model: {best_model_name}")

    joblib.dump(best_model, "resume_classifier.pkl")
    print("✅ Model saved as resume_classifier.pkl")


if __name__ == "__main__":

    path = r"C:\Users\arunr\Desktop\Data Science\Project\Resume Classification\P658_DATASET"

    train_model(path)