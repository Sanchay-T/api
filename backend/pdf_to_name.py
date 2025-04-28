import pdfplumber
import re
import spacy
from .utils import get_base_dir
import os
import logging

logger = logging.getLogger(__name__)
BASE_DIR = get_base_dir()
logger.info("Base Dir : ", BASE_DIR)

def clean_line(line: str) -> str:
    """Clean a single line of text while preserving important information."""
    line = re.sub(r"^\s*(?:\d+\.|\(\d+\)|\d+)\s*", "", line)
    line = re.sub(r"\(cid:\d+\)", " ", line)
    line = line.replace("??", " ")
    line = "".join(char if char.isprintable() else " " for char in line)
    line = re.sub(r"\s+", " ", line)
    return line.strip()


def extract_entities(pdf_path: str, num_lines: int = 10):
    """Extract entities from PDF file and return only the entity texts."""
    try:
        model_path  = os.path.join(BASE_DIR,"models","trained_model_lg_v2_final")
        # print(model_path)
        # model_path  = os.path.join("E:/Workplace/Bizpedia/ats_pyqt/cyphersol-ats-native-app" , "src","utils","trained_model_lg_v2_final")
        nlp = spacy.load(model_path)

        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() or ""

            lines = [clean_line(line) for line in text.split("\n") if clean_line(line)]
            processed_text = "\n".join(lines[:num_lines])
            print("processed_text", processed_text)

            doc = nlp(processed_text)
            entities = [
                ent.text for ent in doc.ents if ent.text
            ]  # Only collect entity texts
            return entities if entities else None

    except Exception as e:
        return None


# Example usage
if __name__ == "__main__":
    pdf_path = "C:/Users/qures/Downloads/Dawat Axis Bank/Dawat Axis Bank/01-08-2020 To  31-07-2021.pdf"
    entities = extract_entities(pdf_path)
    print("Entities:", entities)
    if entities:
        for entity in entities:
            print(f"Detected entity: {entity}")
