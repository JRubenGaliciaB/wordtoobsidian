import os
import re
from docx import Document
from markdownify import markdownify as md
from collections import Counter
import spacy

# Load spaCy model for Named Entity Recognition
nlp = spacy.load("es_core_news_sm")

# Define the input and output directories
input_dir = "columns_word"
output_dir = "columns_markdown"

# Ensure output directory exists
os.makedirs(output_dir, exist_ok=True)

# Relevant topics for public opinion
RELEVANT_TOPICS = {"tecnología", "gobierno", "democracia", "seguridad", "ciudades", "Querétaro", "México", "política",
                   "urbano", "inteligencia artificial", "participación ciudadana"}


# Function to extract keywords and topics using Named Entity Recognition (NER)
def extract_keywords(text, num_keywords=10):
    doc = nlp(text)
    # Extract named entities and filter by type
    entities = [ent.text for ent in doc.ents if ent.label_ in {"PER", "ORG", "LOC", "MISC"}]

    # Include relevant topics explicitly mentioned in the text
    topics = [topic for topic in RELEVANT_TOPICS if topic.lower() in text.lower()]

    # Combine entities and topics, and count frequency
    combined = entities + topics
    word_counts = Counter(combined)

    return [word for word, _ in word_counts.most_common(num_keywords)]


# Process each Word document
for filename in os.listdir(input_dir):
    if filename.endswith(".docx"):
        filepath = os.path.join(input_dir, filename)
        doc = Document(filepath)

        # Combine all text from the Word document
        full_text = "\n".join([p.text for p in doc.paragraphs])

        # Convert to Markdown
        markdown_content = md(full_text)

        # Extract keywords using NER and relevant topics
        keywords = extract_keywords(full_text)

        # Add Obsidian-style links for keywords
        for keyword in keywords:
            markdown_content = re.sub(
                rf"\b({re.escape(keyword)})\b",
                r"[[\1]]",
                markdown_content,
                flags=re.IGNORECASE
            )

        # Save to Markdown file
        output_filename = os.path.splitext(filename)[0] + ".md"
        output_path = os.path.join(output_dir, output_filename)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(markdown_content)

print("Conversion complete! Markdown files with links are saved in the output directory.")
