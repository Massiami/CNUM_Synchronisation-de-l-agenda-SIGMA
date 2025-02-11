# CNUM Projet calendrier

## Objectif 
À partir d'un calendrier de la formation SIGMA au format excel, automatiser l'export de ce calendrier dans différents formats :

* un fichier csv qui contient une ligne par évènement ;
* un fichier ical ou ics qui peut ensuite s'importer dans n'importe quel gestionnaire d'agenda ;
* un agenda google en ligne.

*La difficulté réside dans l'extraction des informations de l'agenda initial.*

Certaines sont à peu près standardisées :

* **Lieu** (ENSAT / UT2J) -> couleur de la cellule
* **Jour et Horaire** -> position de la cellule
* **UE concernée** -> contenu de la cellule

D'autres informations ne sont pas présentes systématiquement et pas de manière standardisées :

* **horaire précis** -> dans les commentaires
* **nom de l'intervenant·e** -> dans les commentaires
* **salle** : -> dans les commentaires

## Livrable
Un dépôt Git qui contient l'ensemble des développements documentés avec une notice d'utilisation.

Voici les différentes fonctionnalités qui peuvent être incrémentées :


*Idéalement, le niveau "Optimal 2" devra être atteint.*

Il est possible que le format initial soit trop contraignant pour atteindre toutes les foncitonnalités, ou peut être pas de manière pérenne. Dans ce cas, les étudiant·es pourront suggérer des modifications du format initial pour faciliter l'automatisation de l'export. Un minimum de modifications sera apprécié par l'équipe enseignante.

## Bibliothèques à utiliser
* Pour lire des fichiers [excel](https://openpyxl.readthedocs.io/en/stable/)
    *permet de charger/écrire, de modifier le style et les commentaires, et bien plus encore. Il faut parcourir la doc depuis le table des matières à gauche et cliquer sur Next Topic.*
    
* Pour exporter au format [ics](https://icspy.readthedocs.io/en/stable/)
    
* Pour exporter dans un google agenda :
    - [Api Python de google](https://developers.google.com/calendar/api/quickstart/python?hl=fr)
    - [un wrapper de l'api google](https://google-calendar-simple-api.readthedocs.io/en/latest/getting_started.html.) 
    *Ça requiert malgré tout la configuration d'un compte google avec des identifiants pour accéder au calendrier.*
    
*Les bibliothèques à utiliser le sont à titre informatif. Si une meilleure solution est trouvée, ainsi soit-t-il.*

## Détails

### Extraction des informations

Voici un aperçu des étapes à réaliser pour faire l'extraction des informations :

0. Ouvrir le fichier et la bonne feuille
1. Parcourir les cellules
2. Déduire la date et l'horaire en fonction de la position de la cellue
3. Obtenir le contenu
4. Déduire le nom de l'UE.
5. Déduire la taille du créneau horaire (2h ou 4h) en 6.fonction de sa position au sein de la cellule
7. Déduire localisation (UT2J / ENSAT) en fonction de la couleur
8. Extraire la description (contient nom intervenant, salle et horaire précise)

*Plusieurs filtres peuvent être rajoutés au cours de cette étape.*

### Formatage des informations et export
#### Format csv

Avec une ligne par évènement et les colonnes suivantes :

* Subject
* Start Date
* Start Time
* End Date
* End Time
* Location
* Description
* Intervenant·es

*Idéalement le nom des colonnes peut être changé facilement dans le code pour s'adapter à des évolutions des standars.*

#### Format ical ou ics
*Voir la documentation des bibliothèques utilisées.*

#### Format google
*Voir la documentation des bibliothèques utilisées. L'agenda doit être public, en lecture seule.*
