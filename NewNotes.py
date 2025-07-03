from docx import Document

def remplacer_identifiants(docx_path, txt_path, output_docx_path):
    # Charger le document Word
    doc = Document(docx_path)

    # Lire les associations identifiant-valeur depuis le fichier texte
    associations = {}
    with open(txt_path, 'r', encoding='utf-8') as file:
        for line in file:
            if '::' in line:
                identifiant, valeur = line.strip().split('::', 1)
                associations[identifiant] = valeur

    # Parcourir chaque paragraphe du document
    for paragraph in doc.paragraphs:
        for identifiant, valeur in associations.items():
            if f"{{{{{identifiant}}}}}" in paragraph.text:
                paragraph.text = paragraph.text.replace(f"{{{{{identifiant}}}}}", valeur)

    # Parcourir chaque tableau du document
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for identifiant, valeur in associations.items():
                    if f"{{{{{identifiant}}}}}" in cell.text:
                        cell.text = cell.text.replace(f"{{{{{identifiant}}}}}", valeur)

    # Sauvegarder le document modifié
    doc.save(output_docx_path)

# Chemins des fichiers
docx_path = "Template_mod.docx"
txt_path = 'notes.txt'
output_docx_path = 'result.docx'
reference = "reference.xlsx"


# Appeler la fonction
remplacer_identifiants(docx_path, txt_path, output_docx_path)
