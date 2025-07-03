import pandas as pd
import re

def df_to_string(dataframe):
    """Transforme un DataFrame en chaîne avec les règles demandées, en nettoyant les doublons."""
    lignes = []
    for _, row in dataframe.iterrows():
        contenu = [str(cell).strip() for cell in row if pd.notna(cell)]
        if contenu:
            ligne = ', '.join(contenu)
            ligne = re.sub(r"Lot\s*\d+\s*:", "", ligne, flags=re.IGNORECASE)  # supprime "Lot X :"
            lignes.append(ligne)
    result = '; '.join(lignes)
    result = re.sub(r'(;\s*){2,}', '; ', result)  # remplace ";;" ou plus par "; "
    return result.strip()

def extraire_tous_les_lots_depuis_excel(fichier_excel, nom_feuille='Devis'):
    df = pd.read_excel(fichier_excel, sheet_name=nom_feuille)

    # Recherche des lignes contenant "Lot X"
    lot_rows = df[df.iloc[:, 0].astype(str).str.contains(r"Lot\s*\d+", case=False, na=False)]
    lot_indices = lot_rows.index.tolist()
    lot_labels = df.iloc[lot_indices, 0].tolist()
    lot_titres = [re.sub(r"Lot\s*\d+\s*:", "", label).strip() for label in lot_labels]

    resultats = {}

    for i, (lot_index, titre) in enumerate(zip(lot_indices, lot_titres)):
        lot_num = i + 1
        fin_index = lot_indices[i+1] if i + 1 < len(lot_indices) else df.shape[0]
        lot_data = df.iloc[lot_index+1:fin_index, 1:3]  # colonnes B et C uniquement

        titre_cle = f"TITRE_AXE{lot_num}"
        contenu_cle = f"AXE{lot_num}"

        resultats[titre_cle] = titre
        resultats[contenu_cle] = df_to_string(lot_data)

    return resultats

# Exemple d'utilisation
if __name__ == "__main__":
    chemin_fichier = "Gjoa MdP/Exemple 2/20250401-Devis-Client-Exemple 2.xlsx"
    resultats = extraire_tous_les_lots_depuis_excel(chemin_fichier)

    for cle, valeur in resultats.items():
        print(f"{cle}:\n{valeur}\n")

def Axes():
    return extraire_tous_les_lots_depuis_excel(chemin_fichier)