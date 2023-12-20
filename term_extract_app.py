from flask import Flask, request, jsonify
from docx import Document
from collections import Counter
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
import string
import io
from flask_cors import CORS  # Import CORS

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

def process_docx(file_content):
    try:
        doc = Document(io.BytesIO(file_content))
        text = " ".join(paragraph.text for paragraph in doc.paragraphs)

        # Tokenize the text
        tokens = word_tokenize(text)

        # Remove punctuation and stopwords
        stop_words = set(stopwords.words('english') + list(string.punctuation))
        filtered_tokens = [word.lower() for word in tokens if word.lower() not in stop_words]

        # Extract key terms using Counter
        term_counter = Counter(filtered_tokens)
        key_terms = [term for term, count in term_counter.most_common(10)]

        return key_terms
    except Exception as e:
        return str(e)

@app.route('/extract-terms', methods=['POST'])
def extract_terms():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'})

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'})

    try:
        content = file.read()
        key_terms = process_docx(content)
        return jsonify({'key_terms': key_terms})
    except Exception as e:
        return jsonify({'error': str(e)})

if __name__ == '__main__':
    app.run(debug=True, port=5000)
