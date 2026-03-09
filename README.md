
# Yabon

## Description

L'objectif de ce projet est de concaténer les données de production hebdomadaire d'une usine d'agroalimentaire dans un tableau annuel. Puis analyser ces données pour trouver les problématiques récurrentes.

Le code n'est pas forcément optimisé, mais a pour objectif de pouvoir être compris et modifié par des personnes qui ne sont pas habituées à la programmation python. De plus, le code ne sera exécuté qu'une seule fois par an donc le coup du programme n'est pas si important.

## Contenu/utilisation

Ce code s'organise en 2 programmes distincts, il est important d'exécuter d'abord "concaténation" puis "analyse_ph".
Il est aussi fortement conseillé de vérifier que les fichiers/dossiers d'entrée/sortie soient bien existants et que leurs noms correspondent, les fichiers de sortie doivent être créés en amont, mais vides.
Les bibliothèques nécessaires au bon fonctionnement du programme sont données dans **Instalation**.

### Concaténation
Ce programme permet de réunir tous les tableaux (de type tableur) dans un seul et unique tableau annuel, tout en complétant les cases vides, uniformisant le texte et retirant les lignes inexploitables. Le programme prend en entrée un dossier source (modifiable à la ligne 8) contenant l'intégralité des fichiers hebdomadaires numéroté de 1 à 53 (modifiable à la ligne 112). Puis rend un fichier sorti (modifiable à la ligne 7). 

"concaténation" utilise aussi un fichier produits.py qui contient un dictionnaire qui prend pour clé le nom de la recette et qui lui associe un code base, un code produit et un format. Il est possible d'en ajouter de nouvelle si nécessaire en respectant la typographie suivante :

'NOM RECETTE': {'code base': '00000', 'code produit': '000000', 'format': 'nom format'}

### Analyse ph
La deuxième partie du programme a pour but d'utiliser en entrée le tableau annuel de l'année précédente (modifiable à la ligne 6), puis d'en extraire les données utiles. Il ressort deux choses :
- Premièrement, dans un tableur (modifiable à ligne 7), le programme créé une feuille par catégorie et donne pour chaque variable de la catégorie le nombre d'erreurs sur le nombre de produits au total. Par exemple pour les batchs à 37 °C, il y aura une ligne pour chaque batch avec : nb non-valides ; nb valides ; nb total.
- Deuxièmement, le programme fournit 6 graphiques de Pareto pour chacune des catégories, mettant en avant les variables qui ont eux le plus d'erreurs. Les graphiques peuvent être copiés-collés en tant qu'image.

## Installation
Pour faire tourner ce programme, vous aurez évidemment besoin d'un logiciel exécutant Python ainsi que des bibliothèques suivantes :
-panda
-matplotlib
-numpy

