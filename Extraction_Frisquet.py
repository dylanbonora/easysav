from datetime import datetime
from pathlib import Path
import sys
import camelot.io as camelot
import fitz
# import sys
from pdf2image import convert_from_path

# RECUPERE LA DATE ET L HEURE DU JOUR
start_time = datetime.now()
print(start_time)

sys.stdout = open('ExtractionTableaux.txt', 'w', encoding='utf-8')

# CREER DOSSIER POUR RECUPERER LES IMAGES TABLEAUX
images_tbl = Path.cwd() / "images"
images_tbl.mkdir(exist_ok=True)

#CREER DOSSIER RAPPORTS
rapport = Path.cwd() / "Rapports"
rapport.mkdir(exist_ok=True)

# LISTE QUI RECUPERERA LES FICHIERS QUI GENERENT DES ERREURS
error_files = []

pdf_files = list((Path.cwd() / 'pdf').glob('*.pdf'))

# BOUCLE SUR LES PDF
for pdf in pdf_files:
    # print('Boucles sur les pdf')
    
    try:
        
        # LISTE QUI RECUPERERA LES DATAFRAMES DU PDF
        dfs_pdf = []
        
        # CREER SOUS-DOSSIER DU MEME NOM QUE LE PDF pour recevoir les images des schemas du pdf
        filename = pdf.stem      # Nom du fichier pdf sans l'extension
        img_pdf_dir = Path.cwd() / 'images' / filename

        if not img_pdf_dir.exists():

            img_pdf_dir.mkdir(exist_ok=True)

        # ********************************************************************
        # CONVERTIR LES PAGES DE SCHEMAS DU PDF EN IMAGES et sauvegarde en jpg
        # ********************************************************************
        
        with fitz.open(pdf) as pages_pdf:
            # print('Ouverture du Dossier PDF')
            
            #  CONVERTIR LES PAGES DU PDF EN IMAGES
            pdf_images = convert_from_path(pdf, 350)
            
            # BOUCLES SUR LES PAGES DU PDF
            for iPage in range(len(pages_pdf)):
                # print('Boucles sur les pages du pdf')

                # On commence par la page 2 des PDF
                if iPage > 2:
                    print(f'page {iPage+1}')
                    
                    # tables_list = camelot.read_pdf(f'pdf/{pdf.name}', flavor='lattice', backend="poppler", split_text=True, line_scale = 100, pages= f'{iPage+1}')
                    tables_list = camelot.read_pdf(f'pdf/{pdf.name}', flavor='stream',pages='all')
                    
                    # LISTE QUI RECUPERERA LES DATAFRAMES DU PDF
                    dfs_pdf = []  # Permet d'afficher l'extraction des tableaux
                    
                    print("Sauvegarde de la page en JPG dans le dossier images ")
                    
                    # SAUVEGARDER LA PAGE EN JPG
                    image_name = f'{str(iPage+1).zfill(2)}_{filename}.jpg'
                    pdf_images[iPage].save(
                    f'images/{filename}/{image_name}', "JPEG")

                    # BOUCLE SUR LA OU LES DATAFRAMES DE LA PAGE
                    for i,table in enumerate(tables_list):
                        
                        if table.df[0][0] == 'Pos':
                            
                            # On ne prend pas en compte les lignes vide du tableaux
                            empty_lines = table.df[table.df[1].map(len) == 0 ].index
                            table.df.drop(empty_lines, inplace=True)
                            
                            # print(table.df[[0,1,2,3,4,5]])
                            print(table.df[[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21]])
             
            # FIN DE LA BOUCLE SUR LES PAGES DU PDF
                                  
    except Exception as err:
        if pdf.stem not in error_files:
            # On ajoute les noms de fichiers 'erreurs' dans la liste error_files
            error_files.append(pdf.stem)

        dfs_pdf = []  # Permet d'afficher l'extraction des tableaux
        print(f'Erreur : {err}') 
    
    
# FIN DE LA BOUCLE SUR LE DOSSIER PDF

# CONVERTIR LA LISTE DES FICHIERS 'ERREURS' EN CHAINE DE CARACTERE pour le fichier rapport
error_files_txt = '\n- '.join(error_files)

# RECUPERER LE NOMBRE DE FICHIERS PDF TRAITES
processfiles_pdf = len(pdf_files)

# CREER UN FICHIER TEXTE DE RAPPORT DE TRAITEMENT
dt = datetime.now()

with open(f"rapports/Extraction_ELM_{dt:%d-%m-%Y_%Hh%Mmn%Ss}.txt", "w+") as report:
    report.write(f"""*********************************
REPORTING du {dt:%d/%m/%Y %H:%M:%S}
*********************************
Nombre de fichiers PDF traites : {processfiles_pdf}

Fichiers non traites : 
- {error_files_txt}""")
    
# Nombre de fichiers CSV en sortie : {processfiles_csv}

# *********************************************************************************************
#  FIN DU SCRIPT
# *********************************************************************************************

# Calcul du temps de traitement :
print('*************************')
time_elapsed = datetime.now() - start_time
print(f'Temps de traitement : (hh:mm:ss.ms)  {time_elapsed}')

# table.rename(columns={0: 'Pos', 1: 'Denomination Descrizioni', 2: 'S-Nr', 3: 'Numero de commande', 4: 'PG', 5: 'AGVA C 21-5M', 6: 'AGVA C 24-5M', 7: 'AGVAC21-6M', 8: 'AGVAC24-6M', 9: 'GVA C 21-5M', 10: 'GVA C 24-5M', 11: 'GVS C 14-5M', 12:' GVS C 24-5M'})