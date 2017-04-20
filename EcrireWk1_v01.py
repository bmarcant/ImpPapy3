#!/usr/bin/python3
# -*- coding: utf8 -*-
"""
==============================================================================
						Ecriture dans un fichier au format WK1 (Lotus 1-2-3)

	Création par B. MARCANT en décembre 2016

==============================================================================
	Modification par ... en ...
		- ....
==============================================================================
"""

import Librairie.Wk1_v01 as Wk1							# Module de gestion du format WK1
import sys																	# Module système (arguments de script)


"""
------------------------------------------------------------------------------
  Variables globales
------------------------------------------------------------------------------
"""
# Constante(s)
# cst_... = ...
cst_Rac = sys.argv[0][0:-3]			# Nom du script (racine)

# Base(s) et table(s)
# bas_... = ...
# tab_... = ...

# Répertoire(s)
# rep_... = ...

# Fichier(s)
# fic_... = ...
fic_Wk1 = cst_Rac + ".wk1"			# Fichier de sortie au format Lotus 1-2-3 WK1

# Liste(s)
# Lst_... = ['A', 'B']

# Dictionnaire(s)
# Dct_... = {'nomA': 'A', 'nomB': 'B']

# Flags
# flg_... = True/False

# Variable(s)
# var_... = ...


"""
------------------------------------------------------------------------------
  Partie principale
------------------------------------------------------------------------------
"""
if __name__ == "__main__" :		# Condition vraie si ce script est exécuté directement dans le shell

	print("Démarrage du script")

	# Instanciation de la classe de gestion du format WK1
	wk1 = Wk1.cl_Wk1()

	# wk1.flgBigEndian = True

	# Création du fichier WK1
	wk1.creer(fic_Wk1)

	# Ecriture d'un texte
	wk1.ecrireTexte("Utilisation de la classe cl_Wk1", 0, 0)

	# Ecriture d'un nombre entier
	wk1.ecrireTexte("Nombre entier court (-32768 à 32767)", 1, 0)
	wk1.ecrireEntier(2016, 1, 1)

	# Ecriture d'un nombre entier long
	wk1.ecrireTexte("Nombre entier long", 2, 0)
	wk1.ecrireEntierLong(9999999, 2, 1)

	# Ecriture d'un nombre décimal avec 2 chiffres après la virgule
	wk1.ecrireTexte("Nombre décimal", 3, 0)
	wk1.ecrireDecimal(1254.28, 3, 1, 2)

	# Ecriture d'un nombre scientifique avec 4 chiffres après la virgule
	wk1.ecrireTexte("Nombre scientifique", 4, 0)
	wk1.ecrireScientifique(1789254.2842, 4, 1, 4)

	# Ecriture d'un nombre monétaire avec 2 chiffres après la virgule
	wk1.ecrireTexte("Nombre monétaire", 5, 0)
	wk1.ecrireMonnaie(45.89, 5, 1, 2)

	# Ecriture d'un pourcentage avec 2 chiffres après la virgule
	wk1.ecrireTexte("Pourcentage", 6, 0)
	wk1.ecrirePourcentage(33.50, 6, 1, 2)

	# Ecriture d'un nombre avec séparateur de milliers et 2 chiffres après la virgule
	wk1.ecrireTexte("Nombre avec séparateur de milliers", 7, 0)
	wk1.ecrireSeparateurMillier(1234345.57, 7, 1, 2)

	# Ecriture d'une date en nombre de jours depuis le 30/12/1899
	# Ecriture d'une date en chaîne de caractères, au format JJ/MM/AAAA
	wk1.ecrireTexte("Date", 8, 0)
	wk1.ecrireDateNombre(42372, 8, 1)
	wk1.ecrireDateChaine("30/12/1899", 8, 2)
	wk1.ecrireDateChaine("04/12/2016", 8, 3)

	# Ecriture d'une heure en nombre de secondes / nombre total de secondes pour une journée
	# Ecriture d'une heure en chaîne de caractères, au format HH:MM:SS
	wk1.ecrireTexte("Heure", 9, 0)
	wk1.ecrireHeureNombre(0.9, 9, 1)
	wk1.ecrireHeureChaine("16:25:37", 9, 2)

	# Ecriture d'un texte vide
	wk1.ecrireTexte("Texte vide", 10, 0)
	wk1.ecrireTexte("", 10, 1)

	# Ecriture de plusieurs cellules successives
	wk1.ecrireTexte("Insertions horizontales", 11, 0)
	wk1.ecrireCellules((2345, 678, 5.23, "Benoît a habité à Oxelaëre."), 11, 1)
	wk1.ecrireTexte("Insertions verticales", 12, 0)
	wk1.ecrireCellules((4.5, 12.3333), 12, 1, 'v')

	# Fermeture du fichier WK1
	wk1.fermer()

	# Destruction de l'objet
	del wk1

	print("Fin du script.")

	exit(0)			# Sortie du script

