ImpPapy3 : convertisseur CSV vers WK1 pour l'import de document dans le logiciel Papyrus2000
============================================================================================

Introduction
------------

Le logiciel [Papyrus2000](http://www.papyrus2000.com) est un très bon logiciel de gestion de bibliothèque, compatible avec toutes les versions de Windows, de XP à Windows 8.

Ce logiciel est désormais disponible en téléchargement gratuit en version complète. 

Hélas, l'incovénient est qu'il n'est plus maintenu depuis plusieurs années.

Et l'un des points bloquant est que Microsoft Excel ne génère plus le format WK1 ([Lotus 1-2-3](https://fr.wikipedia.org/wiki/Lotus_1-2-3)), le format de fichier unique pour l'import dans Papyrus2000.

Bien sûr, il est toujours possible de faire des saisies manuelles pour entrer des documents dans Papyrus2000.

Objet
-----

Aussi, pour palier cette absence, j'ai créé ce script en langage Python qui, à partir d'un fichier au format CSV, génère le fameux fichier au format WK1, indispensable pour l'import dans Papyrus2000.

Utilisation
-----------

Ce script s'utilise en ligne de commande (le terminal sous Linux, `cmd` pour Windows).

Le fichier CSV, en entrée, est celui qui découle de l'export de Papyrus2000 : `ExpPapy3.xls`.

## Linux

Il faut que la version 3 de Python soit installée.

Pour avoir de l'aide :

```bash
$ ImpPapy3_v01.py -h
```

Pour convertir un fichier :

```bash
$ ImpPapy3_v01.py -i fichier.csv -o fichier.wk1
```

## Windows

Utiliser le programme `ImpPapy3.exe` (qui est une compilation du script Python par `cx_Freeze`).

Pour avoir de l'aide :

```bash
C:\ImpPapy3-1.0.exe.win32-3.5\ImpPapy3.exe -h
```

Pour convertir un fichier :

```bash
C:\ImpPapy3-1.0.exe.win32-3.5\ImpPapy3.exe -i fichier.csv -o fichier.wk1
```

Il se peut que les packages [Redistribuable Visual C++](https://support.microsoft.com/en-us/help/2661358/minimum-service-pack-levels-for-microsoft-vc-redistributable-packages) soient nécessaires pour faire fonctionner un exécutable créé avec cx_Freeze. Pour ma part, j'ai installé [Visual C++ Redistributable for Visual Studio 2015](https://www.microsoft.com/fr-FR/download/details.aspx?id=48145).

Version
-------

La version courante est la 1.0.

Note
----

J'ai aussi créé le script `ModPapy3` pour modifier des documents dans Papyrus2000.

