import camelot.io as camelot
from pathlib import Path
import matplotlib.pyplot as plt
import sys
class Logger(object):
    def __init__(self):
        self.terminal = sys.stdout
        self.log = open("stdout.txt", "a")
   
    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)  
    def flush(self):
        # this flush method is needed for python 3 compatibility.
        # this handles the flush command by doing nothing.
        # you might want to specify some extra behavior here.
        pass    
sys.stdout = Logger()
pdf_files = list((Path.cwd() / 'pdf').glob('*.pdf'))
# BOUCLE SUR LA LISTE DES PDF -> Extraction des tableaux PDF vers CSV via Dataframes
for pdf in pdf_files:
  print('**********************************')
  print(pdf.name)
  print('**********************************')
  # EXTRACTION DE TOUTES LES PAGES EN DATAFRAMES
  tables = camelot.read_pdf(f'pdf/{pdf.name}', flavor='stream', pages='all')
  for table in tables:
    # print(table.df)
    if table.df.columns == 4 and table.df[3].apply(lambda x: True if (len(x) < 3) else False):
      print(table.df)
      print('**********************************')
  print('\n\n\n')
  # Exemple de paramètres (De Dietrich) :
  # tables = camelot.read_pdf(f'pdf/{pdf.name}', flavor='stream', table_areas=['0,755,400,0'], columns=['65,106,262,319,367'], pages=f'{iPage+1}').
# flavor = 'stream' -> indique que l'on traite des tableaux sans lignes apparentes
# (par défaut pour des tableaux avec lignes apparentes -> flavor = 'lattice')
# 
# 1ERE PHASE : DETERMINER LES COORDONNEES DU TABLEAU ET (au besoin) DES COLONNES
# table_areas :
# On peut définir une région de la page pour optimiser l'extraction
# (Si on ne veut pas extraire le titre et sous-titres des pages et certaines colonnes inutiles par exemple)
# on définit les coordonnées de table_areas avec matplotlib 
# -> avec tables[n] où n est la première page de tableau du pdf
# (pages indéxées à partir de 0)
# (mais attention, parfois plusieurs mises en page différentes, vérifier les pdf avant)
# columns :
# De la même façon, pour optimiser, on peut indiquer les abscisses des colonnes souhaitées
# (inutile d'indiquer la 1ère abscisse)
# ici on utilise pages='all'
# mais une fois les coordonnées définies,
# on utilisera une variable iPage pour optimiser
# et n'extraire que les pages qui sont des tableaux
    # plt.show() va ouvrir une fenêtre avec un visuel simplifiée de la page
    # permettant avec la souris de récupérer des coordonnées.
    # (Le point d'origine est en bas à gauche)
    # Et table_areas est défini par 2 points : En haut à gauche de la région souhaitée et en bas à droite :
    # table_areas=['x0,y0,x1,y1']
    # (attention : bug avec y1 parfois, laisser à 0)
# table_plot = camelot.plot(tables[1], kind='text')
# plt.show()
# Ensuite, on teste l'extraction et on ajuste au besoin :
# for table in tables:
#   print(table.df)
# 2EME PHASE :
# S'il reste des décalages d'extraction, on les traite à la main
# On vérifie l'extraction de plusieurs pdf pour chercher des bugs d'extraction et les corriger
# On supprime les espaces en trop, les virgules