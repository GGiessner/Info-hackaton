import pandas as pd 
from docx import Document
from generateur_proposition import proposition
from remplissage_axes import remplissage
import docs

DEVIS = docs.devis_3_path
df_devis = pd.read_excel(DEVIS)
REFERENCE = "reference.xlsx"
df_reference = pd.read_excel(REFERENCE)
word = Document("Template.docx")
Contexte_Reference = proposition() #contexte et ref 
Axes_remplis = remplissage() # Titre des axes et sous axes 

def modifie_doc(doc, remplace, remplacant):
    for paragraphe in doc.paragraphs:
        if f"[{remplace}]" in paragraphe.text:
            paragraphe.text = paragraphe.text.replace(remplace, remplacant)

# mets des bullet points 
def modifie_doc_bullet(doc, list):
    

def creation_contrat(doc, dict, doc_sortie):
    for elem in dict.keys():
        modifie_doc(doc, elem, dict[elem])
    doc.save(doc_sortie)

# donne le dictionnaire avec comme clé le nom du lot et comme valeur le prix (colonne 12)

def ajout_dico(df, colonne):
    list = []
    colonne_prix = df.iloc[:, colonne]
    i = 0
    while colonne_prix[i] != "Total arrondi":
        i = i+1
    i = i+1
    for j in range(i, len(colonne_prix)):
        if colonne_prix[j] != None:
            list.append(df_devis.iloc[j, 0] + f" : {colonne_prix[j]}")
    return list

Prix = ajout_dico(df_devis, 12)

for keys, values in Contexte_Reference.items():
    modifie_doc(word, keys, values)
for keys, values in Axes_remplis.items():
    modifie_doc(word, keys, values)
modifie_doc(word, "")