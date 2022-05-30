from tkinter.tix import InputOnly
import camelot
import pandas as pd
from pathlib import Path
import shutil
from datetime import datetime
import fitz  # pip install pymupdf

# RECUPERE LA DATE ET L HEURE DU JOUR
start_time = datetime.now()
print(start_time)

# CREER LE DOSSIER CSV 
csv_dir = Path.cwd() / "csv"
csv_dir.mkdir(exist_ok=True)

# CREER DOSSIER RAPPORTS
report_dir = Path.cwd() / "Rapports"
report_dir.mkdir(exist_ok=True)

# LISTE QUI RECUPERERA LES FICHIERS QUI GENERENT DES ERREURS
error_files = []

# RECUPERER LES FICHIERS PDF A TRAITER DANS UNE LISTE en scannant le dossier PDF
pdf_files = list((Path.cwd() / 'pdf').glob('*.pdf'))

# BOUCLE SUR LA LISTE DES PDF -> Extraction des tableaux PDF vers CSV via Dataframes
for pdf in pdf_files:

    # LISTE QUI RECUPERERA LES DATAFRAMES DU PDF
    dfs_pdf = []

    with fitz.open(pdf) as pages_pdf:

        # BOUCLES SUR LES PAGES DU PDF
        for iPage in range(len(pages_pdf)):

            # CONVERSION DE LA PAGE EN COURS EN LISTE DE DATAFRAMES 
            tables = camelot.read_pdf(f'pdf/{pdf.name}', flavor='stream', table_areas=['0,736,552,62'], pages=f'{iPage+1}')

            # BOUCLE SUR LA OU LES DATAFRAMES DE LA PAGE
            for table in tables:

                # Si plus de 3 colonnes, c'est une page de tableau
                if len(table.df.columns) > 3:

                    try:
                        # NETTOYAGE, REARRANGEMENT DES TABLEAUX

                        # # SUPPRIMER LA COLONNE DES PRIX, colonne index 5
                        table.df.drop(5, axis=1, inplace=True)

                        # Suppression des esapces superflus colonne Designation
                        table.df[2] = table.df[2].str.replace('  ','')

                        # # Renommer les index des colonnes (par défaut : entiers)
                        table.df.rename(columns={0: 'Position', 1: 'Reference', 2: 'Designation', 3: 'GrpMat', 4: 'Quantite'}, inplace=True)

                        # Ajouter colonne Nom du pdf
                        table.df['Modele'] = pdf.stem

                        # Ajouter colonne 'page'
                        table.df['Page'] = f'Page {iPage+1}'

                        # Ajouter colonne 'Substitution' vide (modèles Chappee)
                        table.df['Substitution'] = 'NA'

                        # SUPPRIMER LES LIGNES D'EN TETES
                        # en vérifiant la longueur de la valeur dans la colonne 'Position'
                        # si plus de 4 caractères ou égal à 1, c'est un en-tête 
                        for ligne,value in enumerate(table.df['Position']): 
                            if len(value) > 4 or len(value) == 1:
                                table.df.drop(ligne, inplace=True)   

                        # On reset les index de ligne sinon bug
                        table.df.reset_index(drop=True, inplace=True)

                        # CHERCHER LES LIGNES 'DOUBLES' (texte qui déborde sur une nouvelle ligne)
                        # ET 'REMETTRE' le texte débordant dans la bonne ligne
                        # Puis supprimer la ligne inutile 
                        for ligne,value in enumerate(table.df['Position']): 
                            if value == '':
                                if table.df['Designation'][ligne] != '':
                                    table.df['Designation'][ligne-1] += ' ' + table.df['Designation'][ligne]
                                elif table.df['GrpMat'][ligne] != '':
                                    table.df['GrpMat'][ligne-1] += ' ' + table.df['GrpMat'][ligne]

                                # Supprimer la ligne contenant le texte débordant
                                table.df.drop(ligne, inplace=True) 
         
                        # On reset les index de ligne sinon bug
                        table.df.reset_index(drop=True, inplace=True)

                        # Si valeurs GrpMat décalées dans colonne Designation (bug d'extraction)
                        # Les récupérer et effacer dans Designation
                        for ligne,value in enumerate(table.df['GrpMat']): 
                            if value == '':
                                table.df['GrpMat'][ligne] =  (table.df['Designation'][ligne])[-3:]
                                table.df['Designation'][ligne] = (table.df['Designation'][ligne])[:-3]

                        print(f'pdf : {pdf.stem}')
                        print(table.df)

                        # On ajoute la dataframe à la liste des dataframes du pdf
                        dfs_pdf.append(table.df)

                    except Exception as err:
                        if pdf.stem not in error_files:
                            error_files.append(pdf.stem)  # On ajoute les noms de fichiers 'erreurs' dans la liste error_files

                        dfs_pdf = []

                        print(f'Erreur : {err}')

        # FIN DE LA BOUCLE SUR LES PAGES DU PDF

        # CONCATENER LES DATAFRAMES DU PDF EN UNE SEULE
        if dfs_pdf != [] and pdf.stem not in error_files:
            dfs_pdf = pd.concat(dfs_pdf)

            # CONVERTIR LA DATAFRAME GLOBALE DU PDF EN CSV 
            csv_filepath = Path.cwd() / 'csv' / f'{pdf.stem}.csv'
            dfs_pdf.to_csv(csv_filepath, index=False)

            # SUPPRIMER LES GUILLEMETS ET LES ESPACES SUPERFLUS
            with open(csv_filepath, "r", encoding='utf-8') as text:
                csv_text = text.read().replace('"', '').replace(' ,', ',').replace(', ', ',')

            with open(csv_filepath, "w", encoding='utf-8') as text:
                text.write(csv_text)

# FIN DE LA BOUCLE SUR LE DOSSIER PDF

# CONVERTIR LA LISTE DES FICHIERS 'ERREURS' EN CHAINE DE CARACTERE pour le fichier rapport
error_files_txt = '\n- '.join(error_files)

# RECUPERER LE NOMBRE DE FICHIERS PDF TRAITES
processfiles_pdf = len(pdf_files)

# RECUPERER LE NOMBRE DE FICHIERS CSV TRAITES
csv_files = list((Path.cwd() / "csv").glob('*.csv'))
processfiles_csv = len(csv_files) 

# CREER UN FICHIER TEXTE DE RAPPORT DE TRAITEMENT 
dt = datetime.now()

with open(f"Rapports/Extraction_Viessmann_{dt:%d-%m-%Y_%Hh%Mmn%Ss}.txt", "w+") as report:
    report.write(f"""*********************************
REPORTING du {dt:%d/%m/%Y %H:%M:%S}
*********************************

Nombre de fichiers PDF traites : {processfiles_pdf}
Nombre de fichiers CSV en sortie : {processfiles_csv}

Fichiers non traites : 

- {error_files_txt}""") 

# CONCATENER LES CSV DE CHAQUE PDF EN UN SEUL
with open('csv_final_viessmann.csv', 'w', encoding='utf-8') as outfile:
    for i, fname in enumerate(csv_files):
        with open(fname, 'r', encoding='utf-8') as infile: 
            if i != 0:                  # Supprime les en-têtes sauf celui du 1er csv
                infile.readline()             
            shutil.copyfileobj(infile, outfile)

# FIN DU SCRIPT

# # Calcul du temps de traitement :
print('*************************')
time_elapsed = datetime.now() - start_time
print (f'Temps de traitement : (hh:mm:ss.ms)  {time_elapsed}')

# ----------------------------------------------------------------------------------------


# SNIPPETS pouvant être utiles :

# headlines = table.df[table.df['Position'].map(len) > 4 ].index
# table.df.drop(headlines, inplace=True)    

# # SUPPRIMER LES PRIX (= Garder uniquement les caractères avant l'espace dans colonne 'Quantite') 
# qty = []
# for value in table.df['Quantite']:
#     x = value.find(' ')
#     if x>0:
#         value = value[:x]
#     else:
#         value = value
#     qty.append(value)

# table.df['Quantite'] = qty


# Suppression des guillemets mais pas systématique (si plusieurs guillemets à la suite)
# dfs_pdf.to_csv(csv_filepath, index=False, quoting=csv.QUOTE_NONE,  escapechar="\\")

# df = df.astype('string')

#  AJOUT DES NOMS DE COLONNES dans le csv concaténé
# df = pd.read_csv ('csv_Viessmann.csv', usecols=list(range(0,7)), converters={'0' : str})

# df.to_csv('csv_Viessmann.csv', index=False, header=["Position","Réference","Désignation","GrpMat","Quantité","Modèle","Page"])

# # Find the columns where each value is null (axis=1 -> colonnes)
# cols_vides = [col for col in csv_df.columns if csv_df[col].isnull().all()]
# # Drop these columns from the dataframe
# csv_df.drop(cols_vides, axis=1, inplace=True)

# # Retourne les lignes dont la colonne 0 contient 'Table'
# rows_nan = csv_df[csv_df[0] == 'Table'].index
# # Drop these rows from the dataframe
# csv_df.drop(rows_nan, inplace=True)

# # Retourner une ligne n 
# print(csv_df.iloc[2]) 

# Remplacer une valeur par une autre dans colonne 4
# csv_df[4]= csv_df[4].replace('1 Sur demande' , 'gru')  -> ok

# Supprimer une ou plusieurs colonnes
# csv_df = csv_df.drop(csv_df.columns[4,5], axis=1)

# SUPPRIMER LES COLONNES INUTILES AU DELA DE LA 5EME 
# col_sup4 = [col for col in csv_df if col > 4]
# csv_df.drop(col_sup4, axis=1, inplace=True)

# CONVERTIR LE PDF EN DF
# csv_df = tabula.read_pdf("test.pdf", stream=True, pages=all)
# On obtient une liste de dataframes (une par page du pdf) 
# mais avec des id qui repartent de zéro pour chaque donc non utilisables

# ANCIENNE METHODE AVEC MODULE OS POUR RECUPERER FICHIERS PDF DU DOSSIER COURANT :

    # import os
    # # RECUPERER LES FICHIERS DU DOSSIER COURANT DANS UN TABLEAU
    # files_in_dir = os.listdir(os.path.dirname(__file__))

    # # RECUPERER LE NOM DU FICHIER PDF SANS L'EXTENSION
    # for file in files_in_dir:
    #     if file.endswith('.pdf'):
    #         filename = file.rstrip('.pdf')

# CONVERTIR TOUS LES PDF D'UN DOSSIER
# tabula.convert_into_by_batch("input_directory", output_format='csv', pages='all')

#  SUPPRIMER GUILLEMETS
    # text = open(outfile, "r")
    # text = ' '.join([i for i in text])  
    # text = text.replace('"', '')  
    # textfile = open('textfile.txt', 'w')
    # textfile.write(text)
    # textfile.close()

        # PAGINER LES TABLEAUX ET AJOUTER UNE COLONNE AVEC LE RESULTAT
        # tab_temp = ['Page 1']

        # # SI PDF "Position"
        # if len(csv_df[0][0]) > 4:

        #     page = 1
        #     ligne = 0
        #     for i in range(1, len(csv_df[0])):
        #         ligne += 1    
        #         tab_temp.append('Page ' + str(page))
        #         if len(csv_df[0][i]) > 4:
        #             page += 2
        #             ligne = 0
        #         elif ligne >= 40:
        #             page += 1
        #             ligne = 0
        # #  SI PDF "Article"
        # else:
        #     page = 1
        #     ligne = 0
        #     for i in range(1, len(csv_df[0])):
        #         ligne += 1    
        #         if csv_df[0][i] == '0001':
        #             page += 2
        #             ligne = 0
        #         elif ligne >= 40:
        #             page += 1
        #             ligne = 0
        #         tab_temp.append('Page ' + str(page))
            
        # csv_df[6] = tab_temp

            # # SUPPRIMER LES GUILLEMETS ET AUTRES 
            #     # SUPPRIMER GUILLEMETS ainsi que les espaces avant les virgules (Récupérer les données du csv en chaîne de caractères)
            # with open(csv_filepath, "r", encoding='utf-8') as text:
            #     text = ' '.join([i for i in text])  
            #     text = text.replace('"', '')
            #     text = text.replace(' ,','')

            #     # CREER UN FICHIER TXT avec les données modifiées
            # with open('textfile.txt', 'w', encoding='utf-8') as text_file:
            #     text_file.write(text)

            #     # ECRASER LE CSV AVEC LES DONNEES DU FICHIER TXT
            # with open('textfile.txt', 'r', encoding='utf-8') as in_file:
            #     stripped = (line.strip() for line in in_file)   # Supprime les espaces avant et après
            #     lines = (line.split(",") for line in stripped if line) # Convertit chaque ligne en tableau
            #     with open(csv_filepath, 'w+', newline='', encoding='utf-8') as out_file:
            #         writer = csv.writer(out_file)
            #         writer.writerows(lines)         # Crée le csv ligne par ligne





