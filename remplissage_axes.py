import json
import requests
import sys
from osman import Axes

def parser_sous_axes(texte_axe):
    sous_axes = []
    morceaux = texte_axe.split(";")
    for morceau in morceaux:
        morceau = morceau.strip()
        if "," in morceau:
            nom, desc = morceau.split(",", 1)
            sous_axes.append((nom.strip(), desc.strip()))
        else:
            sous_axes.append((morceau.strip(), None))
    return sous_axes

def generer_texte_mistral(sous_axe, description=None):
    if description:
        prompt = f"Génère une courte phrase en français (peut inclure des mots techniques en anglais) pour décrire succinctement le sous-axe suivant, dans le cadre d'un conseil en stratégie : « {sous_axe} », en t'appuyant sur la description : « {description} »."
    else:
        prompt = f"Génère une courte phrase en français (peut inclure des mots techniques en anglais) pour décrire succinctement le sous-axe suivant, dans le cadre d'un conseil en stratégie : « {sous_axe} »."
    try:
        response = requests.post(
            "http://localhost:11434/api/generate",
            json={"model": "mistral", "prompt": prompt, "stream": False}
        )
        response.raise_for_status()
        return response.json()["response"].strip()
    except Exception as e:
        print(f"Erreur avec le sous-axe '{sous_axe}': {e}")
        return ""

def traiter_axes(Axe):
    resultat = {}
    for i in range(1, len(Axe)//2 + 1):
        titre_cle = f"TITRE_AXE{i}"
        contenu_cle = f"AXE{i}"

        # Copier le titre tel quel
        resultat[titre_cle] = Axe.get(titre_cle, f"Axe {i}")

        # Générer les textes pour chaque sous-axe
        texte = Axe.get(contenu_cle, "")
        sous_axes = parser_sous_axes(texte)
        textes_genérés = []
        for nom, desc in sous_axes:
            texte_genere = generer_texte_mistral(nom, desc)
            textes_genérés.append(texte_genere)

        # Stocker le texte généré dans AXE{i}
        resultat[contenu_cle] = " ; ".join(textes_genérés)
    return resultat


"""
if __name__ == "__main__":
    Axe = Axes()
    nouveau_dictionnaire = traiter_axes(Axe)

    # Affichage correct dans Git Bash
    print(json.dumps(nouveau_dictionnaire, indent=2, ensure_ascii=False), file=sys.stdout)

    # Sauvegarde facultative
    with open("resultat.json", "w", encoding="utf-8") as f:
        json.dump(nouveau_dictionnaire, f, ensure_ascii=False, indent=2)
"""

def remplissage():
    Axe = Axes()
    nouveau_dictionnaire = traiter_axes(Axe)
    return nouveau_dictionnaire

print(remplissage_axes())