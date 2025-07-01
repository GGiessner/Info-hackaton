import pandas as pd 
from docx import Document

DEVIS = "Gjoa MdP/Exemple 1/20250631-Client-Sujet-Devis exemple 1.xlsx"
df_devis = pd.read_excel(DEVIS)
REFERENCE = "reference.xlsx"
df_reference = pd.read_excel(REFERENCE)
dict = {}
excel = ["Lot 1 : Tendances de marché", "Lot 2 : Miroirs clients", "Lot 3 : Concurrence européenne"]
word = Document("Template.docx")

def extraire(df, ligne, colonne):
    dict[f"[2{ligne}"] = df[ligne][colonne]

def modifie_doc(doc, remplace, remplacant):
    for paragraphe in doc.paragraphs:
        if remplace in paragraphe.text:
            paragraphe.text = paragraphe.text.replace(remplace, remplacant)

def creation_contrat(doc, dict, doc_sortie):
    for elem in dict.keys():
        modifie_doc(doc, elem, dict[elem])
    doc.save(doc_sortie)

# cherche les infos dans le devis
for elem in excel:
    extraire(df_devis, elem, 12)

# cherche les expérieneces 

def experience_commune(mot, groupe):
    compteur = 0
    mot_clés = groupe.split()
    for elem in mot_clés:
        if mot == elem:
            return mot 
    return None 

def sujet_commun(mot, df):
    L = []
    for elem in df:
        if experience_commune(mot, elem) != None:
            L.append(experience_commune(mot, elem))
    if len(L) < 8:
        return L
    else:
        return L[:8]     

