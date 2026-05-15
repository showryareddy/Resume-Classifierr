import streamlit as st
import joblib
import re
import nltk

from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer

from docx import Document
import pdfplumber


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

def read_docx(file):
    doc = Document(file)
    return "\n".join([p.text for p in doc.paragraphs])

def read_pdf(file):
    text = ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ""
    return text

@st.cache_resource
def load_model():
    model = joblib.load("model.pkl")
    return model

model = load_model()

st.set_page_config(page_title="Resume Classifier", layout="centered")

st.title("📄 Resume Classification App")
st.write("Upload a resume (PDF or DOCX) to classify it.")

uploaded_file = st.file_uploader(
    "Upload Resume",
    type=["pdf", "docx"]
)

if uploaded_file is not None:

    
    if uploaded_file.name.endswith(".pdf"):
        text = read_pdf(uploaded_file)

    elif uploaded_file.name.endswith(".docx"):
        text = read_docx(uploaded_file)

    else:
        st.error("Unsupported file type")
        st.stop()

   
    cleaned = clean_text(text)


    prediction = model.predict([cleaned])[0]

    st.subheader("✅ Prediction")
    st.success(f"Category: {prediction}")
