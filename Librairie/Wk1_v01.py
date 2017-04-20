#!/usr/bin/python3
# -*- coding: Utf-8 -*-
"""
==============================================================================
					Classes dédiées à la gestion du format WK1 (Lotus 1-2-3)

	Dans un fichier au format WK1, chaque enregistrement est composé au minimum
	d'un numéro d'identification (sur deux octets) et de la longueur des données
	qui suivent (sur deux octets).

	Le premier enregistrement du fichier (0x0000) indique le début du fichier.
	Il contient deux octets de données qui indiquent le numéro de révision du
	fichier Lotus (0x0404 pour Lotus 1-2-3). La longueur des données est donc
	de 0x0002.

	Le dernier enregistrement du fichier (0x0001) indique la fin. Il ne contient
	pas de donnée. La longueur des données est donc de 0x0000.

	Ces deux enregistrements sont les seuls obligatoires du fichier. Entre ces
	deux enregistrements, il peut y avoir n enregistrements.

	Avec LibreOffice, il faut ouvrir un fichier au format WK1 en mode "Western
	Europe (Windows-1252/WinLatin 1)".

	Sources :
		- WSFF1 (Spec introduction).txt
		- WSFF2 (Summary of record types).txt
		- WSFF3 (Cell format encoding).txt

	Création par B. MARCANT en novembre 2016

	Modification par ... en ...
		- ....
==============================================================================
	Note : lancer "python -V" pour avoir la version installée de Python.
==============================================================================
"""

import struct								# Conversion en structures du langage C
# import types								# Interprétation des types de variable (Python 2)
import datetime							# Traitement de date et heure


"""
------------------------------------------------------------------------------
	Classe de gestion du format WK1 (Lotus 1-2-3)

	Pour écrire dans un fichier au format WK1, on utilise les structures du 
	langage C (fonction struct.pack() en Python) :
		"h" -> short integer -> entier court -> 2 octets -> -32768 à 32767
		"d" -> double -> flottant double -> 8 octets -> 1,7x10^308 à 1,7x10^308

	Attribut : - flgBigEndian : mode d'arrangement des octets en mémoire
	
	Méthode  : - __init__ : initialiser l'instance de la classe
	           - __getBigEndian : afficher le mode d'arrangement des octets
	           - __setBigEndian : modifier le mode d'arrangement des octets
	           - creer : créer le fichier au format WK1
						 - ecrireEntier : écrire un nombre entier court
						 - ecrireEntierLong : écrire un nombre entier long
						 - ecrireDecimal : écrire un nombre décimal
						 - ecrireScientifique : écrire un nombre scientifique
						 - ecrireMonnaie : écrire un nombre monétaire
						 - ecrirePourcentage : écrire un pourcentage
						 - ecrireSeparateurMillier : écrire un nombre avec séparateur de milliers
						 - ecrireTexte : écrire un texte
						 - ecrireDateNombre : écrire une date au format nombre de jours
						 - ecrireDateChaine : écrire une date au format chaîne de caractères
						 - ecrireHeureNombre : écrire une heure au format quantième
						 - ecrireHeureChaine : écrire une heure au format chaîne de caractères
						 - ecrirePlage : écrire une plage limitant les cellules définies
	           - ecrireCellules : écrire une ou plusieurs cellules successives
	           - fermer : fermer le fichier ouvert au format WK1
	           - __str__ :  afficher les caractéristiques de l'instance
	           - __del__ :  détruire l'instance de la classe
------------------------------------------------------------------------------
"""
# class cl_Wk1 :					# Version 3 de Python
class cl_Wk1(object) :		# Pour toutes les versions de Python

	# Liste des types d'enregistrement

	__WK1BOF     = 0x0000			# Début de fichier 

														#    Octets | Contenu | Description
														#   --------+---------+-------------
														#    1 - 2  |  00 00  | Type d'enregistrement
														#    3 - 4  |  02 00  | Longueur de données = 2 octets
														#    5 - 6  |  xx xx  | 0x0404 pour fichier Lotus 1-2-3

	__WK1EOF     = 0x0001			# Fin de fichier

														#    Octets | Contenu | Description
														#   --------+---------+-------------
														#    1 - 2  |  01 00  | Type d'enregistrement
														#    3 - 4  |  00 00  | Longueur de données = 0 octet

	__WK1RANGE   = 0x0006			# Cellules minimum et maximum utilisées dans la feuille de calcul

														#    Octets | Contenu | Description
														#   --------+---------+-------------
														#    1 - 2  |  06 00  | Type d'enregistrement
														#    3 - 4  |  08 00  | Longueur de données = 8 octets
														#    5 - 6  |  xx xx  | Colonne en haut à gauche
														#    7 - 8  |  xx xx  | Ligne en haut à gauche
														#    9 - 10 |  xx xx  | Colonne en bas à droite
														#   11 - 12 |  xx 00  | Ligne en bas à droite

	__WK1BLANK   = 0x000C			# Cellule vide pour la définition d'une cellule protégée

														#    Octets | Contenu | Description
														#   --------+---------+-------------
														#    1 - 2  |  0C 00  | Type d'enregistrement
														#    3 - 4  |  05 00  | Longueur de données = 5 octets
														#    5      |  xx     | Format
														#    6 - 7  |  xx xx  | Colonne
														#    8 - 9  |  xx xx  | Ligne

	__WK1INTEGER = 0x000D			# Cellule contenant une valeur entière (short integer)

														#    Octets | Contenu | Description
														#   --------+---------+-------------
														#    1 - 2  |  0D 00  | Type d'enregistrement
														#    3 - 4  |  07 00  | Longueur de données = 7 octets
														#    5      |  xx     | Format
														#    6 - 7  |  xx xx  | Colonne
														#    8 - 9  |  xx xx  | Ligne
														#   10 - 11 |  xx xx  | Valeur

	__WK1NUMBER  = 0x000E			# Cellule contenant une valeur réelle (double float)

														#    Octets | Contenu | Description
														#   --------+---------+-------------
														#    1 - 2  |  0E 00  | Type d'enregistrement
														#    3 - 4  |  0D 00  | Longueur de données = 13 octets
														#    5      |  xx     | Format
														#    6 - 7  |  xx xx  | Colonne
														#    8 - 9  |  xx xx  | Ligne
														#   10 - 11 |  xx xx  | Valeur
														#   12 - 13 |  xx xx  | 
														#   14 - 15 |  xx xx  | 
														#   16 - 17 |  xx xx  | 

	__WK1LABEL   = 0x000F			# Cellule contenant une chaîne de caractères

														#    Octets | Contenu | Description
														#   --------+---------+-------------
														#    1 - 2  |  0F 00  | Type d'enregistrement
														#    3 - 4  |  xx xx  | Longueur de données = (longueur de la chaîne + 7) octets
														#    5      |  xx     | Format
														#    6 - 7  |  xx xx  | Colonne
														#    8 - 9  |  xx xx  | Ligne
														#   10      |  xx     | Cadrage (0x27 pour un justification à gauche)
														#   ...     |  xx xx  | Valeur
														#   ...     |  00     | Fin de la chaîne



	# Liste des formats (ajouter 0x80 pour protéger la cellule)

	__FIXED      = 0x00				# Fixe ("1234"), ajouter 0x0n pour le nombre de décimales
	__SCIENTIFIC = 0x10				# Scientifique ("1,789E+14"), ajouter 0x0n pour le nombre de décimales
	__CURRENCY   = 0x20				# Monnaie ("$45,89"), ajouter 0x0n pour le nombre de décimales
	__PERCENT    = 0x30				# Pourcentage ("33,50%"), ajouter 0x0n pour le nombre de décimales
	__COMMA      = 0x40				# Séparateur de milliers ("34 557,00"), ajouter 0x0n pour le nombre de décimales
	#            = 0x5x				# Non utilisé
	#            = 0x6x				# Non utilisé
	#            = 0x70				# ?
	__LIBRE      = 0x71				# Libre
	__DATEDMY    = 0x72				# Date affichée au format jour mois année ("30 décembre 99")
	__DATEDM     = 0x73				# Date affichée au format jour mois ("30 décembre")
	__DATEMY     = 0x74				# Date affichée au format mois année ("décembre 99")
	__TEXT       = 0x75				# Texte
	__HIDDEN     = 0x76				# Caché (?)
	__TIMEHMS12  = 0x77				# Heure affichée au format heure (12h) minute seconcde AM/PM ("9:36:00 PM")
	__TIMEHM12   = 0x78				# Heure affichée au format heure (12h) minute AM/PM ("9:36 PM")
	__DATEMDY    = 0x79				# Date affichée au format mois jour année ("12/30/99")
	__DATEMD     = 0x7A				# Date affichée au format mois jour ("12/30")
	__TIMEHMS24  = 0x7B				# Heure affichée au format heure (24h) minute seconde ("21:36:00")
	__TIMEHM24   = 0x7C				# Heure affichée au format heure (24h) minute ("21:36")


	"""
	------------------------------------------------------------------------------
		Initialisation de l'instance de la classe
		Argument : flag définissant big ou little endian
		Retour   : néant
	------------------------------------------------------------------------------
	"""
	def __init__(self, flag=False) :
		self.__carEndian = '>' if flag == True else '<'

	"""
	------------------------------------------------------------------------------
		Affichage du mode d'arrangement des octets
		Argument : néant
		Retour   : néant
	------------------------------------------------------------------------------
	"""
	def __getBigEndian(self) :
		return True if self.__carEndian == '>' else False

	"""
	------------------------------------------------------------------------------
		Modification du mode d'arrangement des octets
		Argument : flag définissant big ou little endian
		Retour   : néant

		Selon les OS, les octets ne sont pas ordonnés da la même façon dans un fichier.
			print(struct.pack("<h", 3)) -> b'\x03\x00'
			print(struct.pack(">h", 3)) -> b'\x00\x03'
			print(struct.pack("<i", 3)) -> b'\x03\x00\x00\x00'
			print(struct.pack(">i", 3)) -> b'\x00\x00\x00\x03'
	------------------------------------------------------------------------------
	"""
	def __setBigEndian(self, flag) :
		self.__carEndian = '>' if flag == True else '<'
	
	flgBigEndian = property(__getBigEndian, __setBigEndian)

	"""
	------------------------------------------------------------------------------
		Création du fichier au format WK1
		Argument : nom du fichier à créer
		Retour   : néant
	------------------------------------------------------------------------------
	"""
	def creer(self, ficNom) :
		try :
			self.__ficHdl = open(ficNom, 'wb')		# Ouverture du fichier en écriture binaire
		except IOError :
			print("Erreur : création impossible du fichier \"" + ficNom + "\".")
			exit(1)
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", self.__WK1BOF))				# Type d'enregistrement
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", 0x0002))							# Longueur de données
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", 0x0404))							# Lotus 1-2-3

	"""
	------------------------------------------------------------------------------
		Ecriture d'un nombre entier court (short int)
		Argument : valeur à écrire, ligne et colonne où écrire
		Retour   : néant
	------------------------------------------------------------------------------
	"""
	def ecrireEntier(self, c, lig, col) :
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", self.__WK1INTEGER))		# Type d'enregistrement
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", 0x0007))							# Longueur de données
		self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__FIXED + 0))		# Format
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", col))									# Colonne
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", lig))									# Ligne
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", c))										# Valeur

	"""
	------------------------------------------------------------------------------
		Ecriture d'un nombre entier long (long long int)
		Argument : valeur à écrire, ligne et colonne où écrire
		Retour   : néant
	------------------------------------------------------------------------------
	"""
	def ecrireEntierLong(self, c, lig, col) :
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", self.__WK1NUMBER))		# Type d'enregistrement
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", 0x000D))							# Longueur de données
		self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__FIXED + 0))		# Format
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", col))									# Colonne
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", lig))									# Ligne
		self.__ficHdl.write(struct.pack(self.__carEndian + "d", c))										# Valeur

	"""
	------------------------------------------------------------------------------
		Ecriture d'un nombre décimal (double float)
		Argument : valeur à écrire, ligne et colonne où écrire, nombre de décimals
		Retour   : néant
	------------------------------------------------------------------------------
	"""
	def ecrireDecimal(self, c, lig, col, dec) :
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", self.__WK1NUMBER))		# Type d'enregistrement
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", 0x000D))							# Longueur de données
		self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__FIXED + dec))	# Format
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", col))									# Colonne
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", lig))									# Ligne
		self.__ficHdl.write(struct.pack(self.__carEndian + "d", c))										# Valeur

	"""
	------------------------------------------------------------------------------
		Ecriture d'un nombre scientifique
		Argument : valeur à écrire, ligne et colonne où écrire, nombre de décimals
		Retour   : néant
	------------------------------------------------------------------------------
	"""
	def ecrireScientifique(self, c, lig, col, dec) :
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", self.__WK1NUMBER))		# Type d'enregistrement
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", 0x000D))							# Longueur de données
		self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__SCIENTIFIC + dec))		# Format
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", col))									# Colonne
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", lig))									# Ligne
		self.__ficHdl.write(struct.pack(self.__carEndian + "d", c))										# Valeur

	"""
	------------------------------------------------------------------------------
		Ecriture d'un nombre monétaire
		Argument : valeur à écrire, ligne et colonne où écrire, nombre de décimals
		Retour   : néant
	------------------------------------------------------------------------------
	"""
	def ecrireMonnaie(self, c, lig, col, dec) :
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", self.__WK1NUMBER))		# Type d'enregistrement
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", 0x000D))							# Longueur de données
		self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__CURRENCY + dec))	# Format
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", col))									# Colonne
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", lig))									# Ligne
		self.__ficHdl.write(struct.pack(self.__carEndian + "d", c))										# Valeur

	"""
	------------------------------------------------------------------------------
		Ecriture d'un pourcentage
		Argument : valeur à écrire, ligne et colonne où écrire, nombre de décimals
		Retour   : néant
	------------------------------------------------------------------------------
	"""
	def ecrirePourcentage(self, c, lig, col, dec) :
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", self.__WK1NUMBER))		# Type d'enregistrement
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", 0x000D))							# Longueur de données
		self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__PERCENT + dec))	# Format
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", col))									# Colonne
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", lig))									# Ligne
		self.__ficHdl.write(struct.pack(self.__carEndian + "d", c / 100))							# Valeur

	"""
	------------------------------------------------------------------------------
		Ecriture d'un nombre avec séparateur de milliers
		Argument : valeur à écrire, ligne et colonne où écrire, nombre de décimals
		Retour   : néant
	------------------------------------------------------------------------------
	"""
	def ecrireSeparateurMillier(self, c, lig, col, dec) :
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", self.__WK1NUMBER))		# Type d'enregistrement
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", 0x000D))							# Longueur de données
		self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__COMMA + dec))	# Format
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", col))									# Colonne
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", lig))									# Ligne
		self.__ficHdl.write(struct.pack(self.__carEndian + "d", c))										# Valeur

	"""
	------------------------------------------------------------------------------
		Ecriture d'un texte
		Argument : chaîne de caractères à écrire, ligne et colonne où écrire
		Retour   : néant
	------------------------------------------------------------------------------
	"""
	def ecrireTexte(self, c, lig, col) :
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", self.__WK1LABEL))			# Type d'enregistrement
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", len(c) + 7))					# Longueur de données
		self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__TEXT))					# Format
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", col))									# Colonne
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", lig))									# Ligne
		# self.__ficHdl.write(struct.pack(self.__carEndian + "b", 0x22))								# Justification à droite
		self.__ficHdl.write(struct.pack(self.__carEndian + "b", 0x27))								# Justification à gauche
		# self.__ficHdl.write(struct.pack(self.__carEndian + "b", 0x5E))								# Justification centrée
		self.__ficHdl.write(c.encode("Latin-1"))																			# Valeur
		self.__ficHdl.write(struct.pack(self.__carEndian + "b", 0x00))								# Fin de la chaîne

	"""
	------------------------------------------------------------------------------
		Ecriture d'une date 
		Argument : nombre de jours depuis le 30/12/1899, ligne et colonne où écrire
		Retour   : néant

		Dans les applications Lotus 1-2-3, la date zéro est 30/12/1899.
		Dans les applications Microsoft Office, la date zéro est 12-30-1899.

		C'est le nombre de jours depuis la date zéro de Lotus 1-2-3 qui est écrit en sortie.
	------------------------------------------------------------------------------
	"""
	def ecrireDateNombre(self, c, lig, col) :
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", self.__WK1NUMBER))		# Type d'enregistrement
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", 0x000D))							# Longueur de données
		self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__DATEDMY))			# Format
		# self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__DATEDM))				# Format
		# self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__DATEMY))				# Format
		# self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__DATEMDY))			# Format
		# self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__DATEMD))				# Format
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", col))									# Colonne
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", lig))									# Ligne
		self.__ficHdl.write(struct.pack(self.__carEndian + "d", c))										# Valeur

	"""
	------------------------------------------------------------------------------
		Ecriture d'une date 
		Argument : date au format (JJ/MM/AAAA), ligne et colonne où écrire
		Retour   : néant

		Dans les applications Lotus 1-2-3, la date zéro est 30/12/1899.
		Dans les applications Microsoft Office, la date zéro est 12-30-1899.

		C'est le nombre de jours depuis la date zéro de Lotus 1-2-3 qui est écrit en sortie.
	------------------------------------------------------------------------------
	"""
	def ecrireDateChaine(self, d, lig, col) :
		c = (datetime.datetime.strptime(d, "%d/%m/%Y") - datetime.datetime.strptime("30/12/1899", "%d/%m/%Y")).days
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", self.__WK1NUMBER))		# Type d'enregistrement
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", 0x000D))							# Longueur de données
		self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__DATEDMY))			# Format
		# self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__DATEDM))				# Format
		# self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__DATEMY))				# Format
		# self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__DATEMDY))			# Format
		# self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__DATEMD))				# Format
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", col))									# Colonne
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", lig))									# Ligne
		self.__ficHdl.write(struct.pack(self.__carEndian + "d", c))										# Valeur

	"""
	------------------------------------------------------------------------------
		Ecriture d'une heure
		Argument : quantième de la journée, ligne et colonne où écrire
		Retour   : néant

		C'est le quantième d'une journée de 24 x 60 x 60 = 86400 secondes qui est écrit en sortie.
		Exemple pour 0,9 : 0,9 x 86400 = 77760 secondes => 9:36
	------------------------------------------------------------------------------
	"""
	def ecrireHeureNombre(self, c, lig, col) :
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", self.__WK1NUMBER))		# Type d'enregistrement
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", 0x000D))							# Longueur de données
		# self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__TIMEHMS12))		# Format
		# self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__TIMEHM12))			# Format
		self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__TIMEHMS24))		# Format
		# self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__TIMEHM24))			# Format
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", col))									# Colonne
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", lig))									# Ligne
		self.__ficHdl.write(struct.pack(self.__carEndian + "d", c))										# Valeur

	"""
	------------------------------------------------------------------------------
		Ecriture d'une heure
		Argument : heure au format (HH:MM:SS), ligne et colonne où écrire
		Retour   : néant

		C'est le quantième d'une journée de 24 x 60 x 60 = 86400 secondes qui est écrit en sortie.
		Exemple pour 0,9 : 0,9 x 86400 = 77760 secondes => 9:36
	------------------------------------------------------------------------------
	"""
	def ecrireHeureChaine(self, h, lig, col) :
		c = (datetime.datetime.strptime(h, "%H:%M:%S") - datetime.datetime.strptime("0:00:00", "%H:%M:%S")).seconds / 86400
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", self.__WK1NUMBER))		# Type d'enregistrement
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", 0x000D))							# Longueur de données
		# self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__TIMEHMS12))		# Format
		# self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__TIMEHM12))			# Format
		self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__TIMEHMS24))		# Format
		# self.__ficHdl.write(struct.pack(self.__carEndian + "b", self.__TIMEHM24))			# Format
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", col))									# Colonne
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", lig))									# Ligne
		self.__ficHdl.write(struct.pack(self.__carEndian + "d", c))										# Valeur

	"""
	------------------------------------------------------------------------------
		Ecriture d'une plage limitant les cellules définies
		Argument : ligne et colonne de début, ligne et colonne de fin
		Retour   : néant
	------------------------------------------------------------------------------
	"""
	def ecrirePlage(self, ligdeb, coldeb, ligfin, colfin) :
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", self.__WK1RANGE))			# Type d'enregistrement
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", 0x0008))							# Longueur de données
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", coldeb))							# Colonne en haut à gauche
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", ligdeb))							# Ligne en haut à gauche
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", colfin))							# Colonne en bas à droite
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", ligfin))							# Ligne en bas à droite

	"""
	------------------------------------------------------------------------------
		Ecriture d' une ou plusieurs cellules successives
		Argument : liste de valeurs à écrire, ligne et colonne où écrire
		Retour   : néant
	------------------------------------------------------------------------------
	"""
	def ecrireCellules(self, cellules, lig, col, sens='h') :
		for c in cellules :
			# if type(c) is types.IntType :				# Python 2
			if type(c) is int :										# Python 3
				self.ecrireEntierLong(c, lig, col)
			# elif type(c) is types.FloatType :		# Python 2
			elif type(c) is float :								# Python 3
				self.ecrireDecimal(c, lig, col, 2)
			# elif type(c) is types.StringType :	# Python 2
			elif type(c) is str :									# Python 3
				self.ecrireTexte(c, lig, col)
			else :
				print("Anomalie : type d'enregistrement non géré (" + str(type(c)) + ").")

			if sens in "vV" :
				lig += 1
			else :
				col += 1

	"""
	------------------------------------------------------------------------------
		Fermeture du fichier ouvert au format WK1
		Argument : néant
		Retour   : néant
	------------------------------------------------------------------------------
	"""
	def fermer(self) :
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", self.__WK1EOF))				# Type d'enregistrement
		self.__ficHdl.write(struct.pack(self.__carEndian + "h", 0x0000))							# Longueur de données
		self.__ficHdl.close()		# Fermeture du fichier

	"""
	------------------------------------------------------------------------------
		Affichage des caractéristiques de l'instance
		Argument : néant
		Retour   : néant
	------------------------------------------------------------------------------
	"""
	def __str__(self) :
		return "Gestion du format WK1 (Lotus 1-2-3)."

	"""
	------------------------------------------------------------------------------
		Destruction de l'instance de la classe
		Argument : néant
		Retour   : néant
	------------------------------------------------------------------------------
	"""
	def __del__(self) :
		print("Destruction de l'instance de gestion du format WK1.")


"""
------------------------------------------------------------------------------
  Partie principale
------------------------------------------------------------------------------
"""
if __name__ == "__main__" :		# Condition vraie si ce script est exécuté directement dans le shell

	print("Module de gestion du format WK1 (Lotus 1-2-3).")

	exemple = """
	# Déclaration du module de gestion du format WK1
	import Librairie.Wk1_v01 as Wk1

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
	"""

	print("")
	print("Exemple d'utilisation :")
	print(exemple)

	exit(0)			# Sortie du script

