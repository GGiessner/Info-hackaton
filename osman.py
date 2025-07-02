import pandas as pd
import re

def df_to_string(dataframe):
    """Transforme un DataFrame en chaîne avec les règles demandées."""
    return '; '.join([
        ', '.join([str(cell) for cell in row if pd.notna(cell)])
        for _, row in dataframe.iterrows()
    ])

def extraire_tous_les_lots_depuis_excel(fichier_excel, nom_feuille='Devis'):
    df = pd.read_excel(fichier_excel, sheet_name=nom_feuille)

    # Recherche des lignes contenant "Lot" (Lot 1, Lot 2, etc.)
    lot_rows = df[df.iloc[:, 0].astype(str).str.contains(r"Lot\s*\d+", case=False, na=False)]
    lot_indices = lot_rows.index.tolist()
    lot_labels = df.iloc[lot_indices, 0].tolist()
    lot_titres = [re.sub(r"Lot\s*\d+\s*:", "", label).strip() for label in lot_labels]

    # On prépare le dictionnaire final
    resultats = {}

    for i, lot_index in enumerate(lot_indices):
        lot_num = i + 1
        fin_index = lot_indices[i+1] if i + 1 < len(lot_indices) else df.shape[0]
        lot_data = df.iloc[lot_index+1:fin_index, 1:]

        titre_cle = f"TITRE_AXE{lot_num}"
        contenu_cle = f"AXE{lot_num}"

        resultats[titre_cle] = lot_titres[i]
        resultats[contenu_cle] = df_to_string(lot_data)

    return resultats

# Exemple d'utilisation
if __name__ == "__main__":
    chemin_fichier = "Gjoa MdP/Exemple 2/20250401-Devis-Client-Exemple 2.xlsx"  # adapte le chemin selon ton cas
    resultats = extraire_tous_les_lots_depuis_excel(chemin_fichier)
    for cle, valeur in resultats.items():
        print(f"{cle}:\n{valeur}\n")