# EasySAV

Fonctionnalité qui permet à l'utilisateur d'avoir une infobulle avec diverses informations au toucher du code d'une pièce 
sur le pdf de l'application

-> Reconnaissance de caractères avec le module tesseract de python sur les vues éclatées pour récupérer les coordonnées des pièces (en python)

-> Extraction des données des pièces depuis les pdf fabricants pour générer une base de données (table des pièces) (en python)

-> Récupération des coordonnées du toucher utilisateur (en js)

-> Interrogation de la table des pièces (avec les coord. de la pièce) pour récupérer les infos souhaitées (en php) 

-> Appel de l'API pour récupérer le prix et les infos de livraison (avec le code de la pièce récupéré de la bdd) (php et js)

-> Affichage des infos dans une infobulle (en js)

POUR INSTALLER TESSERACT pour windows :

https://github.com/UB-Mannheim/tesseract/wiki

POUR INSTALLER LE MODULE CAMELOT avec ses dépendances:

https://camelot-py.readthedocs.io/en/master/user/install.html

POUR INSTALLER POPPLER (nécessaire pour le module pdf2image):

https://pdf2image.readthedocs.io/en/latest/installation.html

POUR INSTALLER LES MODULES NECESSAIRES, en ligne de commande :

pip install -r requirements.txt

(Installer Java pour le module tabula-py : https://www.java.com/en/download/manual.jsp)
