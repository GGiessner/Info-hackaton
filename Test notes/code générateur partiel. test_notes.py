import pandas as pd
import re
from docx import Document

## Paramètres: Attention, il faut remplacer le nom des dossier par le nom des fichers de la mission

NOTES_FILE = "test_notes.txt"         # Notes brutes
WORD_TEMPLATE = "template_mod.docx"  # Modèle Word
WORD_OUTPUT = "result.docx"      # Fichier final
doc = Document(WORD_TEMPLATE)

## Infos notes

notes_data = {}
with open(NOTES_FILE, 'r') as f:
    for line in f:
        if "::" in line:
            ident, content = line.strip().split("::", 1)
            notes_data[ident] = content

## code note

for paragraph in doc.paragraphs:
    for match in re.finditer(r'\{1(\w+)\}', paragraph.text):
        ident = match.group(1)
        if ident in notes_data:
            paragraph.text = paragraph.text.replace(match.group(0), notes_data[ident])

# Sauvegarder nouveau Word

doc.save(WORD_OUTPUT)

print(f" Nouveau document créé : {WORD_OUTPUT}")

