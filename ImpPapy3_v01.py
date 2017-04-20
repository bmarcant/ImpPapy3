#!/usr/bin/python3
# -*- coding: utf8 -*-
"""
==============================================================================
						Ecriture dans un fichier au format WK1 (Lotus 1-2-3)
											pour import dans Papyrus2000

	Création par B. MARCANT en décembre 2016

==============================================================================
	Modification par ... en ...
		- ....
==============================================================================
"""

import sys																	# Gestion système (arguments de script)
import getopt																# Interprétation des options
import time																	# Gestion du temps
import datetime							# Traitement de date et heure
import Librairie.Wk1_v01 as Wk1							# Gestion du format WK1
import csv																	# Gestion du format CSV
import re																		# Expressions régulières


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
fic_Log = cst_Rac + ".log"			# Fichier de log
fic_Csv = "ImpPapy3.csv"				# Fichier CSV par défaut (entrée)
# fic_Csv = "/sdcard/qpython/projects3/ImpPapy3/" + fic_Csv				# Fichier CSV sur Androïd
fic_Wk1 = "ImpPapy3.wk1"				# Fichier WK1 par défaut (sortie)

# Liste(s)
# Lst_... = ['A', 'B']

# Dictionnaire(s)
# Dct_... = {'nomA': 'A', 'nomB': 'B']

# Flags
# flg_... = True/False
flg_Vrb = False									# Mode verbeux

# Variable(s)
# var_... = ...


"""
------------------------------------------------------------------------------
  Classe de paramètres pour la définition du fichier CSV
	Attribut : cf. ci-dessous
	Méthode  : néant
------------------------------------------------------------------------------
"""
class cl_Dialect(csv.Dialect) :
	delimiter = ';'									# Séparateur de champ (',' par défaut)
	doublequote = None							# Doublement de quotechar inclus dans les données
	escapechar = '\\'								# Caractère d'échapement pour protéger les caractères spéciaux (None par défaut)
	lineterminator = "\r\n"					# Fin de ligne
	quotechar = '"'									# Délimiteur de chaîne ('"' par défaut)
	quoting = csv.QUOTE_NONE				# Ajout automatique du délimiteur de chaîne (QUOTE_MINIMAL par défaut)
	# quoting = csv.QUOTE_NONNUMERIC	# Ajout automatique du délimiteur de chaîne (QUOTE_MINIMAL par défaut)
	skipinitialspace = True					# Gestion des espaces entre le délimiteur de chaîne et les données


"""
------------------------------------------------------------------------------
  Fonction d'affichage de l'usage de ce script
	Argument : néant
	Retour   : néant
------------------------------------------------------------------------------
"""
def fn_Usage() :
	print("")
	print("Syntaxe : {} [-h] [-i fichier_CSV] [-o fichier_WK1] [-v]".format(sys.argv[0]))
	print("  -h : aide.")
	print("  -i fichier_CSV : fichier CSV en entrée (par défaut : \"{}\").".format(fic_Csv))
	print("  -o fichier_WK1 : fichier WK1 en sortie (par défaut : \"{}\").".format(fic_Wk1))
	print("  -v : mode verbeux.")
	print("")
	print("Les champs du fichier d'entrée ne doivent comporter ni point-virgule, ni retour chariot.")
	print("")
	print("Le fichier d'entrée CSV peut être généré à partir de LibreOffice, sous Linux. Dans ce cas, il faut le générer au format \"Europe occidentale (Windows-1252/WinLatin1)\".")
	print("")


"""
------------------------------------------------------------------------------
  Fonction d'affichage d'un message sur la console et dans un fichier de log
	Argument : message à afficher, mode (lecture, écriture, ajout)
	Retour   : néant
------------------------------------------------------------------------------
"""
def fn_EcrireLog(msg, mode='a') :
	print("")
	print("{}".format(msg))
	ficLog = open(fic_Log, mode)
	# ficLog.write("\n{} : {}\n".format(time.strftime("%d/%m/%Y %H:%M:%S", time.localtime()), msg))
	ficLog.write("\n{}\n".format(msg))
	ficLog.close()


"""
------------------------------------------------------------------------------
  Fonction retournant True si l'argument est un entier
	Argument : chaîne de caractères
	Retour   : True si l'argument est un entier, sinon False
------------------------------------------------------------------------------
"""
def fn_isEntier(chaine) :
  return re.match("^[0-9]+$", chaine) != None


"""
------------------------------------------------------------------------------
  Fonction retournant True si l'argument est un décimal
	Argument : chaîne de caractères
	Retour   : True si l'argument est un décimal, sinon False
------------------------------------------------------------------------------
"""
def fn_isDecimal(chaine) :
  return re.match("^[0-9]*\.[0-9]+$", chaine) != None


"""
------------------------------------------------------------------------------
  Fonction retournant True si l'argument est une date
	Argument : chaîne de caractères
	Retour   : True si l'argument est une date, sinon False
------------------------------------------------------------------------------
"""
def fn_isDate(chaine) :
	# dat = re.match("^([0-9]{1,2})[\/\-\.]([0-9]{1,2})[\/\-\.]([1-2][0-9]{1,3})$", chaine)
	dat = re.match("^([0-9]{2})[\/]([0-9]{2})[\/]([1-2][0-9]{3})$", chaine)		# Format JJ/MM/AAAA uniquement
	if dat != None :
		try :
			datetime.date(int(dat.group(3)), int(dat.group(2)), int(dat.group(1)))
			return True
		except ValueError :
			return False
	else :
		return False


"""
------------------------------------------------------------------------------
  Fonction retournant une chaîne de caractères au format date
	Argument : chaîne de caractères
	Retour   : chaîne de caractères au format date, vide si aucune date détectée
------------------------------------------------------------------------------
"""
def fn_toDate(chaine) :
	# return re.match("^[0-9]{1,2}[\/\-\.][0-9]{1,2}[\/\-\.][1-2][0-9]{1,3}$", chaine) != None
	dat = re.match("^([0-9]{1,2})[\/\-\.]([0-9]{1,2})[\/\-\.]([1-2][0-9]{1,3})$", chaine)
	if dat != None :
		try :
			dt = datetime.date(int(dat.group(3)), int(dat.group(2)), int(dat.group(1)))
			return dt.strftime("%d/%m/%Y")
		except ValueError :
			return ""
	else :
		return ""


"""
------------------------------------------------------------------------------
  Partie principale
------------------------------------------------------------------------------
"""
if __name__ == "__main__" :		# Condition vraie si ce script est exécuté directement dans le shell

	# Création du fichier log
	fn_EcrireLog("Démarrage du script \"{}\" ({}).".format(sys.argv[0], time.strftime("%d/%m/%Y %H:%M:%S", time.localtime())), mode='w')

	# Lecture des options
	try :
		opts, args = getopt.getopt(sys.argv[1:], "hi:o:v", ["help", "input=", "output=", "verbose"])
	except getopt.GetoptError as err :
		fn_EcrireLog("Erreur : {}.".format(err))
		fn_Usage()							# Affichage de l'aide
		sys.exit(2)							# Sortie du script avec erreur

	for opt, arg in opts :
		if opt in ("-h", "--help") :
			fn_Usage()						# Affichage de l'aide
			sys.exit()						# Sortie du script
		elif opt in ("-i", "--input") :
			fic_Csv = arg
		elif opt in ("-o", "--output") :
			fic_Wk1 = arg
		elif opt in ("-v", "--verbose") :
			flg_Vrb = True
		else :
			fn_EcrireLog("Anomalie : option {} inconnue".format(opt))
			# sys.exit(2)

	# Affichage des paramètres
	fn_EcrireLog("Fichier de log : \"{}\".".format(fic_Log))

	if flg_Vrb :
		fn_EcrireLog("Mode verbeux : oui.")
		fn_EcrireLog("Fichier CSV en entrée : \"{}\".".format(fic_Csv))
		fn_EcrireLog("Fichier Wk1 en entrée : \"{}\".".format(fic_Wk1))

	# Ouverture du fichier CSV en entrée
	fn_EcrireLog("Ouverture du fichier d'entrée.")

	ficEnt = None

	try :
		ficEnt = open(fic_Csv, 'r', encoding="Latin-1")		# Ouverture du fichier en lecture
	except IOError :
		fn_EcrireLog("Erreur : ouverture impossible du fichier \"{}\".".format(fic_Csv))
		exit(1)
	
	# Initialisation du dictionnaire des colonnes
	dctColonnes = {}														# Création d'un dictionnaire vide
	dctColonnes['Reference'] = 99999
	dctColonnes['Titre'] = ""
	dctColonnes['Auteur'] = ""
	dctColonnes['Editeur'] = "*"
	dctColonnes['Code Classement'] = "002"
	dctColonnes['Genre'] = "1"
	dctColonnes['XSignet'] = ""
	dctColonnes['Prete a'] = ""
	dctColonnes['CodeLecteur'] = ""
	dctColonnes['Prete le'] = ""
	dctColonnes['Rendu le'] = ""
	dctColonnes['Annee d\'edition'] = ""
	dctColonnes['Date d\'achat'] = ""
	dctColonnes['Origine'] = ""
	dctColonnes['Type de document'] = "1"
	dctColonnes['Cote'] = ""
	dctColonnes['Prix'] = ""
	dctColonnes['Lieu'] = "2"
	dctColonnes['NbPage'] = ""
	dctColonnes['Contenant'] = ""
	dctColonnes['Collection'] = "C.R.G.F.A. Bailleul"
	dctColonnes['Photo1'] = ""
	dctColonnes['Son1'] = ""
	dctColonnes['Frequence'] = ""
	dctColonnes['Duree'] = ""
	dctColonnes['Format'] = ""
	dctColonnes['Annee origine'] = ""
	dctColonnes['ISBN'] = ""
	dctColonnes['ISSN'] = ""
	dctColonnes['Etat'] = ""
	dctColonnes['NbEx'] = ""
	dctColonnes['Appreciation'] = "1"
	dctColonnes['Codebarre'] = "99999"
	dctColonnes['AdrInternet'] = ""
	dctColonnes['Valeur'] = ""
	dctColonnes['Pays Ville'] = ""
	dctColonnes['Tomaison'] = ""
	dctColonnes['Commentaires'] = ""
	dctColonnes['Resume'] = ""
	dctColonnes['NbPret'] = ""
	dctColonnes['Debut abonnement'] = ""
	dctColonnes['Dernier abonnement'] = ""
	dctColonnes['Prochain abonnement'] = ""
	dctColonnes['Periodicite'] = ""
	dctColonnes['Nbr abonnement'] = ""
	dctColonnes['Langue'] = ""
	dctColonnes['Systeme'] = ""
	dctColonnes['Version'] = ""
	dctColonnes['Son'] = ""
	dctColonnes['Pellicule'] = ""
	dctColonnes['Info'] = ""
	dctColonnes['NumNotice'] = ""
	dctColonnes['DateImportNotice'] = ""
	dctColonnes['DateRetourNotice'] = ""
	dctColonnes['NomFichierNotice'] = ""
	dctColonnes['NomOrigineNotice'] = ""
	dctColonnes['Alpha1'] = ""
	dctColonnes['Alpha2'] = ""
	dctColonnes['Alpha3'] = ""
	dctColonnes['Num1'] = ""
	dctColonnes['Num2'] = ""
	dctColonnes['Date1'] = ""
	dctColonnes['Public'] = "Tous publics"

	# Création du fichier WK1 en sortie
	fn_EcrireLog("Création du fichier de sortie.")

	wk1 = Wk1.cl_Wk1()							# Instanciation de la classe de gestion du format WK1

	# wk1.flgBigEndian = True				# Mode Big-Endian ou Little-Endian

	wk1.creer(fic_Wk1)							# Création du fichier WK1

	# Lecture du fichier CSV en entrée, avec titres de colonne
	fn_EcrireLog("Lecture du fichier d'entrée.")
	# lignes = csv.DictReader(ficEnt)														# Dictionnaire
	lignes = csv.DictReader(ficEnt, dialect=cl_Dialect())			# Dictionnaire

	l = 0
	for lig in lignes :
		if flg_Vrb :
			fn_EcrireLog("Ligne {} : Référence = \"{}\", Titre = \"{}\", Editeur = \"{}\", Prêté le = \"{}\"...".format(l + 1, lig['Reference'], lig['Titre'], lig['Editeur'], lig['Prete le']))

		# Initialisation des valeurs par défaut
		col = dctColonnes
		col['Reference'] = str(l + 1)

		# Parcours des index de la ligne
		for cle in lig.keys() :
			if cle == "Reference" :
				col['Reference'] = lig['Reference']
			elif cle == "Titre" :
				col['Titre'] = lig['Titre']
			elif cle == "Auteur" :
				col['Auteur'] = lig['Auteur']
			elif cle == "Editeur" :
				col['Editeur'] = lig['Editeur']
			elif cle == "Code Classement" :
				col['Code Classement'] = lig['Code Classement']
			elif cle == "Genre" :
				col['Genre'] = lig['Genre']
			elif cle == "XSignet" :
				col['XSignet'] = lig['XSignet']
			elif cle == "Prete a" :
				col['Prete a'] = lig['Prete a']
			elif cle == "CodeLecteur" :
				col['CodeLecteur'] = lig['CodeLecteur']
			elif cle == "Prete le" :
				col['Prete le'] = lig['Prete le']
			elif cle == "Rendu le" :
				col['Rendu le'] = lig['Rendu le']
			elif cle == "Annee d'edition" :
				col['Annee d\'edition'] = lig['Annee d\'edition']
			elif cle == "Date d'achat" :
				col['Date d\'achat'] = lig['Date d\'achat']
			elif cle == "Origine" :
				col['Origine'] = lig['Origine']
			elif cle == "Type de document" :
				col['Type de document'] = lig['Type de document']
			elif cle == "Cote" :
				col['Cote'] = lig['Cote']
			elif cle == "Prix" :
				col['Prix'] = lig['Prix']
			elif cle == "Lieu" :
				col['Lieu'] = lig['Lieu']
			elif cle == "NbPage" :
				col['NbPage'] = lig['NbPage']
			elif cle == "Contenant" :
				col['Contenant'] = lig['Contenant']
			elif cle == "Collection" :
				col['Collection'] = lig['Collection']
			elif cle == "Photo1" :
				col['Photo1'] = lig['Photo1']
			elif cle == "Son1" :
				col['Son1'] = lig['Son1']
			elif cle == "Frequence" :
				col['Frequence'] = lig['Frequence']
			elif cle == "Duree" :
				col['Duree'] = lig['Duree']
			elif cle == "Format" :
				col['Format'] = lig['Format']
			elif cle == "Annee origine" :
				col['Annee origine'] = lig['Annee origine']
			elif cle == "ISBN" :
				col['ISBN'] = lig['ISBN']
			elif cle == "ISSN" :
				col['ISSN'] = lig['ISSN']
			elif cle == "Etat" :
				col['Etat'] = lig['Etat']
			elif cle == "NbEx" :
				col['NbEx'] = lig['NbEx']
			elif cle == "Appreciation" :
				col['Appreciation'] = lig['Appreciation']
			elif cle == "Codebarre" :
				col['Codebarre'] = lig['Codebarre']
			elif cle == "AdrInternet" :
				col['AdrInternet'] = lig['AdrInternet']
			elif cle == "Valeur" :
				col['Valeur'] = lig['Valeur']
			elif cle == "Pays Ville" :
				col['Pays Ville'] = lig['Pays Ville']
			elif cle == "Tomaison" :
				col['Tomaison'] = lig['Tomaison']
			elif cle == "Commentaires" :
				col['Commentaires'] = lig['Commentaires']
			elif cle == "Resume" :
				col['Resume'] = lig['Resume']
			elif cle == "NbPret" :
				col['NbPret'] = lig['NbPret']
			elif cle == "Debut abonnement" :
				col['Debut abonnement'] = lig['Debut abonnement']
			elif cle == "Dernier abonnement" :
				col['Dernier abonnement'] = lig['Dernier abonnement']
			elif cle == "Prochain abonnement" :
				col['Prochain abonnement'] = lig['Prochain abonnement']
			elif cle == "Periodicite" :
				col['Periodicite'] = lig['Periodicite']
			elif cle == "Nbr abonnement" :
				col['Nbr abonnement'] = lig['Nbr abonnement']
			elif cle == "Langue" :
				col['Langue'] = lig['Langue']
			elif cle == "Systeme" :
				col['Systeme'] = lig['Systeme']
			elif cle == "Version" :
				col['Version'] = lig['Version']
			elif cle == "Son" :
				col['Son'] = lig['Son']
			elif cle == "Pellicule" :
				col['Pellicule'] = lig['Pellicule']
			elif cle == "Info" :
				col['Info'] = lig['Info']
			elif cle == "NumNotice" :
				col['NumNotice'] = lig['NumNotice']
			elif cle == "DateImportNotice" :
				col['DateImportNotice'] = lig['DateImportNotice']
			elif cle == "DateRetourNotice" :
				col['DateRetourNotice'] = lig['DateRetourNotice']
			elif cle == "NomFichierNotice" :
				col['NomFichierNotice'] = lig['NomFichierNotice']
			elif cle == "NomOrigineNotice" :
				col['NomOrigineNotice'] = lig['NomOrigineNotice']
			elif cle == "Alpha1" :
				col['Alpha1'] = lig['Alpha1']
			elif cle == "Alpha2" :
				col['Alpha2'] = lig['Alpha2']
			elif cle == "Alpha3" :
				col['Alpha3'] = lig['Alpha3']
			elif cle == "Num1" :
				col['Num1'] = lig['Num1']
			elif cle == "Num2" :
				col['Num2'] = lig['Num2']
			elif cle == "Date1" :
				col['Date1'] = lig['Date1']
			elif cle == "Public" :
				col['Public'] = lig['Public']
			elif cle != None :
				fn_EcrireLog("Anomalie : index \"{}\" inconnu.".format(cle))
			else :
				pass

		# Ecriture des valeurs dans le fichier WK1
		if fn_isEntier(col['Reference']) :
			wk1.ecrireEntierLong(int(col['Reference']), l, 0)
		else :
			wk1.ecrireTexte(col['Reference'], l, 0)

		wk1.ecrireTexte(col['Titre'], l, 1)

		wk1.ecrireTexte(col['Auteur'], l, 2)
		
		wk1.ecrireTexte(col['Editeur'], l, 3)

		wk1.ecrireTexte(col['Code Classement'], l, 4)

		if fn_isEntier(col['Genre']) :
			wk1.ecrireEntier(int(col['Genre']), l, 5)
		else :
			wk1.ecrireTexte(col['Genre'], l, 5)

		if fn_isEntier(col['XSignet']) :
			wk1.ecrireEntier(int(col['XSignet']), l, 6)
		else :
			wk1.ecrireTexte(col['XSignet'], l, 6)

		wk1.ecrireTexte(col['Prete a'], l, 7)

		if fn_isEntier(col['CodeLecteur']) :
			wk1.ecrireEntierLong(int(col['CodeLecteur']), l, 8)
		else :
			wk1.ecrireTexte(col['CodeLecteur'], l, 8)

		if fn_toDate(col['Prete le']) != "" :
			wk1.ecrireDateChaine(col['Prete le'], l, 9)
		elif fn_isEntier(col['Prete le']) :
			wk1.ecrireDateNombre(int(col['Prete le']), l, 9)
		else :
			wk1.ecrireTexte(col['Prete le'], l, 9)

		if fn_toDate(col['Rendu le']) != "" :
			wk1.ecrireDateChaine(col['Rendu le'], l, 10)
		elif fn_isEntier(col['Rendu le']) :
			wk1.ecrireDateNombre(int(col['Rendu le']), l, 10)
		else :
			wk1.ecrireTexte(col['Rendu le'], l, 10)

		if fn_isEntier(col['Annee d\'edition']) :
			wk1.ecrireEntier(int(col['Annee d\'edition']), l, 11)
		else :
			wk1.ecrireTexte(col['Annee d\'edition'], l, 11)

		if fn_toDate(col['Date d\'achat']) != "" :
			wk1.ecrireDateChaine(col['Date d\'achat'], l, 12)
		elif fn_isEntier(col['Date d\'achat']) :
			wk1.ecrireDateNombre(int(col['Date d\'achat']), l, 12)
		else :
			wk1.ecrireTexte(col['Date d\'achat'], l, 12)

		wk1.ecrireTexte(col['Origine'], l, 13)

		if fn_isEntier(col['Type de document']) :
			wk1.ecrireEntier(int(col['Type de document']), l, 14)
		else :
			wk1.ecrireTexte(col['Type de document'], l, 14)

		if fn_isEntier(col['Cote']) :
			wk1.ecrireEntierLong(int(col['Cote']), l, 15)
		else :
			wk1.ecrireTexte(col['Cote'], l, 15)

		if fn_isEntier(col['Prix']) or fn_isDecimal(col['Prix']) :
			wk1.ecrireDecimal(int(col['Prix']), l, 16, 2)
		else :
			wk1.ecrireTexte(col['Prix'], l, 16)

		if fn_isEntier(col['Lieu']) :
			wk1.ecrireEntierLong(int(col['Lieu']), l, 17)
		else :
			wk1.ecrireTexte(col['Lieu'], l, 17)

		if fn_isEntier(col['NbPage']) :
			wk1.ecrireEntier(int(col['NbPage']), l, 18)
		else :
			wk1.ecrireTexte(col['NbPage'], l, 18)

		if fn_isEntier(col['Contenant']) :
			wk1.ecrireEntierLong(int(col['Contenant']), l, 19)
		else :
			wk1.ecrireTexte(col['Contenant'], l, 19)

		wk1.ecrireTexte(col['Collection'], l, 20)

		wk1.ecrireTexte(col['Photo1'], l, 21)

		wk1.ecrireTexte(col['Son1'], l, 22)

		if fn_isEntier(col['Frequence']) :
			wk1.ecrireEntierLong(int(col['Frequence']), l, 23)
		else :
			wk1.ecrireTexte(col['Frequence'], l, 23)

		if fn_isEntier(col['Duree']) :
			wk1.ecrireEntierLong(int(col['Duree']), l, 24)
		else :
			wk1.ecrireTexte(col['Duree'], l, 24)

		wk1.ecrireTexte(col['Format'], l, 25)

		if fn_isEntier(col['Annee origine']) :
			wk1.ecrireEntier(int(col['Annee origine']), l, 26)
		else :
			wk1.ecrireTexte(col['Annee origine'], l, 26)

		wk1.ecrireTexte(col['ISBN'], l, 27)

		wk1.ecrireTexte(col['ISSN'], l, 28)

		wk1.ecrireTexte(col['Etat'], l, 29)

		if fn_isEntier(col['NbEx']) :
			wk1.ecrireEntier(int(col['NbEx']), l, 30)
		else :
			wk1.ecrireTexte(col['NbEx'], l, 30)

		wk1.ecrireTexte(col['Appreciation'], l, 31)

		if fn_isEntier(col['Codebarre']) :
			wk1.ecrireEntierLong(int(col['Codebarre']), l, 32)
		else :
			wk1.ecrireTexte(col['Codebarre'], l, 32)

		wk1.ecrireTexte(col['AdrInternet'], l, 33)

		if fn_isEntier(col['Valeur']) :
			wk1.ecrireEntierLong(int(col['Valeur']), l, 34)
		else :
			wk1.ecrireTexte(col['Valeur'], l, 34)

		wk1.ecrireTexte(col['Pays Ville'], l, 35)

		wk1.ecrireTexte(col['Tomaison'], l, 36)

		wk1.ecrireTexte(col['Commentaires'], l, 37)

		wk1.ecrireTexte(col['Resume'], l, 38)

		if fn_isEntier(col['NbPret']) :
			wk1.ecrireEntier(int(col['NbPret']), l, 39)
		else :
			wk1.ecrireTexte(col['NbPret'], l, 39)

		if fn_toDate(col['Debut abonnement']) != "" :
			wk1.ecrireDateChaine(col['Debut abonnement'], l, 40)
		elif fn_isEntier(col['Debut abonnement']) :
			wk1.ecrireDateNombre(int(col['Debut abonnement']), l, 40)
		else :
			wk1.ecrireTexte(col['Debut abonnement'], l, 40)

		if fn_toDate(col['Dernier abonnement']) != "" :
			wk1.ecrireDateChaine(col['Dernier abonnement'], l, 41)
		elif fn_isEntier(col['Dernier abonnement']) :
			wk1.ecrireDateNombre(int(col['Dernier abonnement']), l, 41)
		else :
			wk1.ecrireTexte(col['Dernier abonnement'], l, 41)

		if fn_toDate(col['Prochain abonnement']) != "" :
			wk1.ecrireDateChaine(col['Prochain abonnement'], l, 42)
		elif fn_isEntier(col['Prochain abonnement']) :
			wk1.ecrireDateNombre(int(col['Prochain abonnement']), l, 42)
		else :
			wk1.ecrireTexte(col['Prochain abonnement'], l, 42)

		wk1.ecrireTexte(col['Periodicite'], l, 43)

		if fn_isEntier(col['Nbr abonnement']) :
			wk1.ecrireEntier(int(col['Nbr abonnement']), l, 44)
		else :
			wk1.ecrireTexte(col['Nbr abonnement'], l, 44)

		if fn_isEntier(col['Langue']) :
			wk1.ecrireEntier(int(col['Langue']), l, 45)
		else :
			wk1.ecrireTexte(col['Langue'], l, 45)

		wk1.ecrireTexte(col['Systeme'], l, 46)

		wk1.ecrireTexte(col['Version'], l, 47)

		wk1.ecrireTexte(col['Son'], l, 48)

		wk1.ecrireTexte(col['Pellicule'], l, 49)

		wk1.ecrireTexte(col['Info'], l, 50)

		wk1.ecrireTexte(col['NumNotice'], l, 51)

		if fn_toDate(col['DateImportNotice']) != "" :
			wk1.ecrireDateChaine(col['DateImportNotice'], l, 52)
		elif fn_isEntier(col['DateImportNotice']) :
			wk1.ecrireDateNombre(int(col['DateImportNotice']), l, 52)
		else :
			wk1.ecrireTexte(col['DateImportNotice'], l, 52)

		if fn_toDate(col['DateRetourNotice']) != "" :
			wk1.ecrireDateChaine(col['DateRetourNotice'], l, 53)
		elif fn_isEntier(col['DateRetourNotice']) :
			wk1.ecrireDateNombre(int(col['DateRetourNotice']), l, 53)
		else :
			wk1.ecrireTexte(col['DateRetourNotice'], l, 53)

		wk1.ecrireTexte(col['NomFichierNotice'], l, 54)

		wk1.ecrireTexte(col['NomOrigineNotice'], l, 55)

		wk1.ecrireTexte(col['Alpha1'], l, 56)

		wk1.ecrireTexte(col['Alpha2'], l, 57)

		wk1.ecrireTexte(col['Alpha3'], l, 58)

		wk1.ecrireTexte(col['Num1'], l, 59)

		wk1.ecrireTexte(col['Num2'], l, 60)

		if fn_toDate(col['Date1']) != "" :
			wk1.ecrireDateChaine(col['Date1'], l, 61)
		elif fn_isEntier(col['Date1']) :
			wk1.ecrireDateNombre(int(col['Date1']), l, 61)
		else :
			wk1.ecrireTexte(col['Date1'], l, 61)

		wk1.ecrireTexte(col['Public'], l, 62)

		l += 1

	# Ecriture de la plage dans le fichier  WK1
	wk1.ecrirePlage(0, 0, l, len(dctColonnes))

	# Fermeture du fichier en entrée
	fn_EcrireLog("Fermeture du fichier en entrée.")
	ficEnt.close()

	# Fermeture du fichier en sortie
	fn_EcrireLog("Fermeture du fichier en sortie.")
	wk1.fermer()
	del wk1										# Destruction de l'objet

	# Fin du script
	fn_EcrireLog("Fin du script \"{}\" ({}).".format(sys.argv[0], time.strftime("%d/%m/%Y %H:%M:%S", time.localtime())))

	exit(0)			# Sortie du script

