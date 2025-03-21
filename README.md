# README CNUM_Synchronisation de l'agenda SIGMA

## I- Quelques Rappels concernant le projet :

### Contexte :
Le Master Géomatique SIGMA a pour objectif de former aux métiers de l’environnement et de l’aménagement impliquant la maîtrise de concepts, méthodes et techniques liés à la gestion de l’information géographique. 
Les cours de ce master ont lieu à deux endroits différents : l’ENSAT et l’université Jean Jaurès (UT2J). Ces établissements utilisent des systèmes différents pour communiquer l’emploi du temps aux étudiants. Ainsi, depuis de nombreuses années, les professeurs doivent remplir un fichier Excel où sont répertoriés différentes informations comme :
- le lieu du cours (ENSAT/UT2J)
- le jour et l’horaire
- l’UE concernée
- le nom de l’intervenant
- la salle. 

Cependant, les étudiants doivent systématiquement télécharger en ligne la dernière version du fichier Excel pour consulter leur emploi du temps, ce qui n’est pas optimal. 

### Objectifs du projet :
Permettre aux étudiants d’avoir accès à un agenda en ligne contenant toutes les informations initialement présentes dans l’Excel. 
Ne pas changer les habitudes des professeurs qui utilisent cet agenda Excel depuis de nombreuses années. 
Pour cela : Automatiser le formatage du calendrier de cours du master SIGMA du format Excel vers un format csv pour ensuite le publier en ligne sous forme d’un Google Agenda. 

### Résultat :
Nous sommes parvenus à générer un code Python qui permet de passer du fichier Excel de base à un Google Agenda. Nous y récupérons toutes les informations à savoir : la salle du cours, l’UE concernée, le jour et l’horaire standard, les informations des commentaires (nom du professeur intervenant et horaires précis). 
Afin que ce script fonctionne correctement, différentes choses doivent être faites au préalable. Nous détaillons cela dans la suite de ce document. 

## II- A faire en amont

Dans ce dépôt GitHub, vous trouverez différents fichiers. Afin que le programme fonctionne correctement, vous devez :
1. Créer un dossier dans votre environnement de travail et mettre le classeur Excel de base où vous ferez les modifications. 
2. Télécharger les documents suivants et les mettre dans ce même dossier : 
   - `credentials.json`
   - `token.json`
   - `config.txt`
   - `CNUM_SIGMA.py`
3. Remplir le fichier `config.txt` qui contiendra les chemins d’accès à ces documents. Tout est indiqué dans ce fichier afin de vous aider à le remplir correctement. 
4. Ouvrir la console du système de votre ordinateur ou anaconda prompt si vous avez Spider. Puis copier-coller ceci : 
   - pip install google-api-python-client
   - pip install google-auth-oauthlib
5. Assurez-vous d’avoir l’identifiant et le mot de passe de l’adresse mail du master SIGMA. 
6. Sur le calendrier Excel de base, s’assurer que : 
- La colonne date est bien la colonne E
- La colonne Vendredi après-midi est bien la colonne O
- La ligne avec “Lu Matin”, “..., “Ve Aprem” est bien la ligne 5
- La ligne correspondant à la semaine 11 est la ligne 33

Normalement, aucun changement à faire, ce sont les paramètres par défaut de votre classeur Excel.

## III- Lorsque vous faites des modifications d’emplois du temps sur le classeur Excel 

1. Faites vos modifications de façon classique.
2. Enregistrez.  
3. Ouvrez le fichier `CNUM_SIGMA.py` sur votre environnement de travail (Thonny, Pyzo, Spider, Python, etc.).
4. Exécutez le script. 
5. Ouvrez l’agenda Google associé à l’adresse mail du master SIGMA : vous avez accès à l’agenda. 

## IV- Accès des étudiants à l’agenda

Afin que les étudiants puissent avoir accès à l’agenda Google en ligne, nous pensons que le mieux est un partage d’agenda Google au début de l’année selon la procédure suivante : 

1. Un professeur ouvre Google Agenda.
2. À gauche, cliquez sur le nom du nouvel agenda.
3. Placez le curseur sur l'agenda partagé et cliquez sur Plus Paramètres et partage.
4. Sélectionnez une option :
- Tous les membres de votre organisation : sous "Autorisations d'accès", cochez "Rendre disponible pour votre organisation". 
- Pour partager un agenda, demandez aux utilisateurs de s'y abonner, ou partagez-le avec une personne ou un groupe.
- Dans la zone des autorisations, cliquez sur la flèche vers le bas et choisissez une option. Pour en savoir plus, consultez Paramètres d'autorisation.
5. Cliquez sur Envoyer.
6. Les invitations à des agendas groupés incluent des liens vers les agendas.
7. Les utilisateurs reçoivent des notifications par e-mail lorsque des agendas sont partagés. Ces notifications par e-mail contiennent un lien Ajouter à l'agenda. Si un utilisateur clique sur ce lien, puis sur Ajouter un agenda, l'agenda s'affiche dans la liste "Autres agendas" de cet utilisateur.

## V- Pistes d’amélioration

Afin d’optimiser davantage ce code, nous avons identifié différentes pistes d’améliorations possibles ainsi que quelques recommandations : 

1. Arriver à récupérer les horaires précis dans les commentaires et les ajuster directement sur l’agenda. Pour l’instant, le code parvient uniquement à adapter l’horaire quand il détecte un commentaire au format “9h-12h”. Par exemple, “8h30-12h30” ne fonctionne pas. De plus, nous recommandons une harmonisation lors de l’écriture des nouveaux horaires. Nous pensons que le mieux est le format “9h-12h” à la ligne, sans rien d’autre avant ou après. 
2. Afin d'harmoniser la présentation des commentaires, nous recommandons aux professeurs d'utiliser le format suivant :
- Horaire précis : [ex. 10h-12h]
- Nom de l’intervenant(e) : [ex. M. Marc Lang]
- Salle de cours : [ex. 1113 Ensat]
- Autres informations : [ex. prévoir un ordinateur portable]

Chaque commentaire devra respecter ce format afin d’assurer une meilleure lisibilité et organisation des informations.  

3. Suivre les modifications faites par les professeurs avec la création d’une feuille au sein du classeur Excel qui recense le avant/après pour avoir un suivi des modifications.
4. Envoyer un mail à chaque personne utilisant l’agenda pour indiquer qu’une modification (avec le nom de l’UE et la date) a été effectuée.
5. Avoir la possibilité de filtrer l’emploi du temps par UE, Lieux ou Intervenant. Pour cela, utiliser la fonction `groupby` de pandas.
