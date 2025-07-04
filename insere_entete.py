import pandas as pd 
from docx import Document
from generateur_proposition import proposition
from remplissage_axes import remplissage
import docs
from docx.oxml import OxmlElement

def remplacer_phrase_entete(doc_path, ancienne_phrase, nouvelle_phrase, output_path):
    doc = Document(doc_path)
    for section in doc.sections:
        header = section.header
        for para in header.paragraphs:
            if ancienne_phrase in para.text:
                para.text = para.text.replace(ancienne_phrase, nouvelle_phrase)
    doc.save(output_path)

remplacer_phrase_entete(docs.template_path, )
