import os
import re
import json
import requests
import pandas as pd
from docxtpl import DocxTemplate

# === PARAMÈTRES ===
CHEMIN_NOTES = "Gjoa MdP/Exemple 1/20250631-Gjoa-Sujet-Proposition.docx"
CHEMIN_DEVIS = "Gjoa MdP/Exemple 1/20250631-Client-Sujet-Devis exemple 1.xlsx"
CHEMIN_MODELE = "Gjoa MdP/Exemple 1/20250631-Gjoa-Sujet-Proposition.docx"
NOM_SORTIE = "proposition_generee.docx"

def lire_docx(path):
    from docx import Document
    doc = Document(path)
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]

def lire_excel(path):
    try:
        return pd.read_excel(path)
    except:
        return pd.DataFrame()

def extraire_sections(paragraphs):
    sections = {}
    current_key = None
    for line in paragraphs:
        titre = re.match(r"^(contexte|besoins|objectifs|problèmes|livrables|planning|tarif|budget|contact|interlocuteur)[\s:]*$", line.strip().lower())
        if titre:
            current_key = titre.group(1)
            sections[current_key] = ""
        elif current_key:
            sections[current_key] += line.strip() + " "
    return sections

def construire_prompt(sections, devis_df):
    prompt = "Tu es un consultant expérimenté.\n"
    prompt += "Rédige une proposition d'intervention professionnelle à insérer dans un modèle Word. Pour chaque section, fournis un bloc clair comme :\n"
    prompt += "## Contexte\ntexte\n## Objectifs\ntexte\n## Livrables\ntexte\n..."

    for k, v in sections.items():
        prompt += f"\n\n{k.capitalize()} : {v.strip()}"

    if not devis_df.empty:
        prompt += "\n\nPrestations prévues (issues du devis) :\n"
        for _, row in devis_df.iterrows():
            ligne = " - " + ", ".join(str(val) for val in row if pd.notnull(val))
            prompt += ligne + "\n"

    return prompt

def requete_mistral(prompt):
    texte_complet = ""
    response = requests.post(
        "http://localhost:11434/api/chat",
        json={
            "model": "mistral",
            "messages": [{"role": "user", "content": prompt}],
            "stream": True
        }
    )
    for ligne in response.iter_lines():
        if ligne:
            data = json.loads(ligne.decode('utf-8'))
            if 'message' in data and 'content' in data['message']:
                texte_complet += data['message']['content']
    return texte_complet

def parser_blocs(texte):
    # Extrait les blocs type ## Titre\nContenu
    blocs = {}
    sections = re.split(r"^##\s+", texte, flags=re.MULTILINE)
    for bloc in sections[1:]:
        titre_ligne, *contenu = bloc.strip().split("\n", 1)
        if contenu:
            blocs[titre_ligne.strip().lower()] = contenu[0].strip()
    return blocs

def remplir_modele(blocs, fichier_modele, fichier_sortie):
    doc = DocxTemplate(fichier_modele)
    doc.render(blocs)
    doc.save(fichier_sortie)

# === MAIN ===
if __name__ == "__main__":
    if not all(os.path.exists(p) for p in [CHEMIN_NOTES, CHEMIN_DEVIS, CHEMIN_MODELE]):
        exit(1)

    notes = lire_docx(CHEMIN_NOTES)
    devis = lire_excel(CHEMIN_DEVIS)
    sections = extraire_sections(notes)

    prompt = construire_prompt(sections, devis)
    texte_genere = requete_mistral(prompt)
    blocs = parser_blocs(texte_genere)
    remplir_modele(blocs, CHEMIN_MODELE, NOM_SORTIE)
