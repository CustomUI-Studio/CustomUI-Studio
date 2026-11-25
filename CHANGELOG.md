# Historique des versions

La dernière version se télécharge sur la page [Releases](https://github.com/CustomUI-Studio/CustomUI-Studio/releases/latest).

## 1.20.1

Correctif de l'aperçu en direct : il montre le ruban comme Excel. Les projets et les classeurs des versions précédentes s'ouvrent et s'enregistrent comme avant.

### Corrections

- **Grands boutons** : le libellé s'écrit sur deux lignes équilibrées, comme dans Excel (« Nouvel / Article », « Envoyer par / courriel ▾ »), et le bouton prend la largeur de la plus longue ligne, sans « … ». Même règle pour les groupes d'Excel dessinés dans l'aperçu.
- **Libellé** (`labelControl`) : texte seul, sans image.
- **Contrôle sans image** (bouton, bascule, menu, bouton fractionné, galerie) : aucune image, comme dans Excel ; en grande taille, l'emplacement de l'icône reste vide.
- **Menu ou galerie avec une imageMso** : son image est affichée.
- **Case à cocher** : décochée au départ, comme dans Excel.
- **Icône manquante** : une image générique remplace l'icône dans l'aperçu, la structure, l'inspecteur et les catalogues.
- **Structure du ruban** : les éléments du menu Fichier (Backstage) et chaque menu clic droit se replient.
- **Panneau Propriétés** : son en-tête reste sur une ligne, avec le type de l'élément en majuscules (traduit en anglais) ; replié, toute la bande le rouvre, même en bas d'une longue page ; ouvert, un clic sur sa barre de titre le replie.

## 1.20.0

Raccourcis clavier et IA installée sur l'ordinateur. Les projets et les classeurs des versions précédentes s'ouvrent et s'enregistrent comme avant.

### Nouveautés

- **Raccourcis clavier** (panneau Callbacks VBA, section « Raccourcis clavier ») : désactiver une touche d'Excel (Ctrl+C, Ctrl+V, F2…) ou la confier à une procédure du module `RibbonDESIGNER`, par `Application.OnKey`. Les touches sont posées à l'ouverture et quand le classeur revient au premier plan, et rendues à Excel quand il perd la main ou se ferme (extrait `ThisWorkbook` à coller). Le préréglage « Sans copier-coller » du clic droit peut désactiver aussi les six raccourcis du copier-coller (Ctrl+C, Ctrl+X, Ctrl+V, Ctrl+Inser, Maj+Inser, Maj+Suppr). Vérifié dans Excel pour Microsoft 365.
- **Assistant IA : mode « IA sur cet ordinateur »** (version de bureau, facultatif) : la demande affichée part en un clic vers Ollama ou LM Studio installés sur votre ordinateur, et la réponse est vérifiée automatiquement avant toute modification. Seules les adresses de votre ordinateur sont acceptées : rien n'est envoyé sur Internet. Le copier-coller reste le mode par défaut.

### À savoir

- Raccourcis : sans effet pendant la saisie dans une cellule et dans les boîtes de dialogue ; ce n'est pas une protection. Dans un complément, fermer le complément sans quitter Excel laisse les touches prises jusqu'au redémarrage d'Excel : appeler d'abord `RestaurerRaccourcis` (exemple dans le module).
- IA sur cet ordinateur : installer soi-même Ollama ou LM Studio et un modèle (gratuits) ; un modèle d'au moins 7 milliards de paramètres est conseillé. Sur le processeur, une demande peut prendre plusieurs minutes.
- Il n'y a pas de version 1.19.0 : la 1.20.0 réunit les phases 19 et 20.

## 1.18.0

Construire un ruban de bout en bout depuis l'application. Les projets et les classeurs des versions précédentes s'ouvrent et s'enregistrent comme avant.

### Nouveautés

- **Panneau « Problèmes » dans le Constructeur visuel**, sous l'aperçu : erreurs et avertissements du ruban pendant la construction, avec les règles de l'enregistrement. Un clic sélectionne l'élément concerné ; un problème sans élément ouvre l'Éditeur XML sur la bonne ligne. Pastilles dans l'arbre et encadré « Problèmes de cet élément » dans l'inspecteur.
- **Dupliquer, copier et coller** un onglet, un groupe ou un contrôle (barre de l'inspecteur, Ctrl+D, Ctrl+C, Ctrl+V), aussi d'un projet à l'autre. Identifiants libres (`<id>_2`) ; les procédures VBA dérivées de l'identifiant sont renommées, les procédures partagées restent communes ; le raccourci clavier (`keytip`) n'est pas copié.
- **« Appliquer à un nouveau classeur… »** : classeur vierge (`.xlsm` si le ruban appelle des macros, sinon `.xlsx`) qui reçoit le ruban du projet.
- **Guide « Terminer dans Excel »** après chaque « Appliquer », rouvrable depuis le panneau Callbacks VBA : seulement les étapes utiles au projet. Dans l'application de bureau, modules VBA enregistrés à côté du classeur (jamais remplacés sans confirmation) et bouton « Ouvrir dans Excel ».

### À savoir

- CustomUI Studio ne modifie toujours jamais le code VBA d'un classeur : les modules s'importent dans Excel.
- Dans un `.xlsm` neuf, tant que les modules ne sont pas importés, Excel signale à l'ouverture qu'il ne peut pas exécuter la macro `CustomUI_OnLoad` : cliquer OK, puis importer les modules.

## 1.15.0

Un champ pour chaque attribut utile du schéma de Microsoft. Les projets et les classeurs des versions précédentes s'ouvrent et s'enregistrent comme avant.

### Nouveautés

- **Onglets** (personnalisés et contextuels) : visibilité (`visible`), callbacks `getVisible`, `getLabel` et `getKeytip`, étiquette (`tag`). Le libellé et la visibilité se pilotent aussi par une macro avec `SetControlState`, comme pour un contrôle.
- **Groupes** : icône du groupe replié (`imageMso` ou image du classeur), réduction automatique des grands boutons quand la fenêtre rétrécit (`autoScale`), centrage vertical (`centerVertically`), visibilité et étiquette. Libellé et visibilité pilotables par `SetControlState`.
- **Bouton d'un lanceur de boîte de dialogue** : état actif, callbacks (`getEnabled`, `getVisible`…), raccourci clavier, affichage et étiquette.
- **Galerie et liste modifiable** : actualisation à chaque ouverture (`invalidateContentOnDrop`) ; galerie : `showInRibbon` ; séparateur de menu : titre.
- **Réaffectation des commandes d'Office** (`<commands>`) à l'écran : nouvelle section « Commandes Office (réaffectation) » ; une commande d'Excel peut être grisée (`enabled`, `getEnabled`) ou remplacée par une macro quand elle est lancée depuis le ruban (`onAction`).
- **« Autres attributs du schéma »** : tout attribut permis à cet endroit sans champ dédié reste saisissable.

### Corrections

- Les champs que le schéma ne permet pas à cet endroit ne sont plus proposés (saisie perdue jusqu'ici) : libellé du bouton fractionné (il se règle sur son bouton), idMso, visibilité et image du lanceur de boîte de dialogue, des séparateurs et des conteneurs.
- Un onglet personnalisé n'apparaît plus barré dans l'arbre quand `startFromScratch` vaut `true`.
- Un contrôle Office sans `idMso` bloque l'enregistrement : Excel l'ignorait en affichant « ID manquant ».

## 1.14.1

Correctif de l'aperçu du Backstage et du format d'un classeur ouvert. Les projets et les classeurs ne changent pas.

### Corrections

- **Aperçu du Backstage.** Les colonnes ne se chevauchent plus : comme dans Excel, une colonne s'élargit selon son contenu, la deuxième commence après la première et la page défile horizontalement si besoin.
- **Format du classeur ouvert.** À l'ouverture d'un classeur `.xlsm` ou `.xlam`, son format devient celui du projet : la liste « Format » l'affiche (grisée, « — classeur ouvert ») et, pour un complément `.xlam`, le module RibbonDESIGNER est généré en mode complément. Un classeur `.xlsx` ne change pas le format du projet.

### Amélioration

- **Panneau Propriétés repliable** pour laisser plus de place à l'aperçu ; son état est mémorisé. Marges réduites autour de l'aperçu.

## 1.14.0

Réglages d'affichage du classeur générés en VBA. Les projets et les classeurs des versions précédentes s'ouvrent et s'enregistrent comme avant.

### Nouveautés

- **Affichage du classeur.** Nouvelle section du panneau « Callbacks VBA » : barre de formule, barre d'état, ruban entier, onglets de feuilles, barres de défilement, en-têtes, quadrillage (« ne pas toucher », « masquer » ou « afficher ») et titre de la fenêtre. La portée de chaque réglage est indiquée : tout Excel, ce classeur ou toutes les feuilles. Rien n'est écrit dans le XML du ruban.
- **Code généré.** Le module RibbonDESIGNER applique les réglages à l'ouverture (`AppliquerAffichage`) et rend à Excel ce qui agit sur tout Excel quand un autre classeur devient actif ou que le classeur se ferme (`RestaurerAffichage`). Un bouton copie l'extrait à coller dans ThisWorkbook ; avec l'option CommandBars, un seul extrait couvre les deux.
- **Complément (.xlam).** Seuls les réglages qui agissent sur tout Excel sont proposés (un complément n'a pas de fenêtre).
- Comportements relevés dans Excel pour Microsoft 365, puis vérifiés sur un classeur construit par l'application.

## 1.13.0

Tous les groupes d'Excel, et un aperçu qui les dessine comme Excel. Les projets et les classeurs des versions précédentes s'ouvrent et s'enregistrent comme avant.

### Nouveautés

- **Tous les groupes d'Excel.** La liste « Copier un groupe d'Excel dans cet onglet » propose maintenant les groupes de tous les onglets (Insertion, Formules, Données, Affichage…), classés par onglet dans l'ordre d'Excel, puis les autres groupes documentés par Microsoft pour la version d'Excel visée. Les positions « avant / après » proposent tous les groupes de l'onglet.
- **Aperçu fidèle.** Les groupes d'Excel sont dessinés avec leurs vrais boutons, listes et cases à cocher, à leur place et à leur taille ; dans l'application de bureau, avec les icônes d'Excel. Le bouton « Alt » affiche aussi les raccourcis d'Excel.
- Ruban relevé dans Excel pour Microsoft 365 ; chacun des groupes proposés comme vérifiés a été copié dans un onglet et affiché par Excel.

## 1.12.0

Onglets et groupes d'Excel dans l'éditeur visuel, raccourcis clavier et réglages d'affichage. Les projets et les classeurs des versions précédentes s'ouvrent et s'enregistrent comme avant.

### Nouveautés

- **Onglets et groupes d'Excel.** Ajoutez un onglet d'Excel (Accueil, Insertion…) à votre projet pour y placer vos groupes, le masquer, ou masquer certains de ses groupes d'une case. Copiez un groupe d'Excel, par exemple Police, dans votre propre onglet. Avec `startFromScratch`, l'onglet d'Excel ajouté est réaffiché (Excel l'exige).
- **Position.** Placez votre onglet avant ou après un onglet d'Excel (par exemple juste après Accueil), et vos groupes avant ou après un groupe d'Excel du même onglet.
- **Raccourcis clavier.** Champ `keytip` sur les onglets, les groupes et les contrôles, avec un avertissement quand une lettre est déjà prise par Excel (Excel la remplacerait sans rien dire).
- **Réglages d'affichage.** Nouvelle section de l'inspecteur : grisé, image seule ou libellé seul, largeur et longueur des zones de saisie, disposition des boîtes, grandes entrées de menu avec description, colonnes de galerie… Seuls les réglages que le schéma permet pour l'élément choisi sont proposés.
- **Aperçu.** Les onglets d'Excel à leur place, leurs groupes par leur nom, vos réglages appliqués, et un bouton « Alt » qui affiche les raccourcis.
- Comportements relevés dans Excel pour Microsoft 365 et vérifiés sur un classeur réel.

## 1.11.1

Correctif de l'aperçu de la barre d'outils Accès rapide. Les projets et les classeurs ne changent pas.

### Corrections

- **Aperçu de la barre d'outils Accès rapide.** Les boutons avec une image personnalisée l'affichent (seules les icônes intégrées d'Office apparaissaient). Comme dans Excel, un bouton sans libellé n'affiche que son image ; son nom reste visible en info-bulle.

### Amélioration

- **Structure du ruban.** Les blocs « Barre d'outils Accès rapide », « Menus contextuels », « Onglets contextuels » et « Backstage » se replient avec un petit « − » et se déplient avec « + ». « Tout replier » et « Tout déplier » les incluent.

## 1.11.0

Barre d'outils Accès rapide dans l'éditeur visuel. Les projets et les classeurs des versions précédentes s'ouvrent et s'enregistrent comme avant.

### Nouveautés

- **Barre d'outils Accès rapide.** Nouvelle section de la Structure du ruban : ajoutez des commandes d'Excel (catalogue avec badge de compatibilité), vos propres boutons reliés à une procédure VBA et des séparateurs, puis réordonnez-les. Deux listes : « Ce classeur », affichée quand le classeur est actif, et « Partagée », affichée dans tous les classeurs pour un complément `.xlam`.
- **Ce qu'Excel exige, expliqué.** Excel n'applique la barre que si `startFromScratch` vaut « true » (sinon il l'ignore et affiche un message d'erreur), et la barre du classeur remplace celle de l'utilisateur. Au premier ajout, l'application propose de garder Annuler et Refaire et de passer `startFromScratch` à « true », en rappelant que les onglets d'Excel sont alors masqués. Rien ne change sans votre choix.
- **Aperçu.** La barre s'affiche sous le ruban, avec ses icônes et ses libellés, comme dans Excel.
- Comportements relevés dans Excel pour Microsoft 365 et vérifiés sur un classeur réel.

## 1.10.0

Assistant IA par cible : ruban, onglet contextuel ou menu Fichier (Backstage). Les projets et les classeurs des versions précédentes s'ouvrent et s'enregistrent comme avant.

### Nouveautés

- **Choix de la cible.** Avant de décrire ce que vous voulez, choisissez « Ruban », « Onglet contextuel » ou « Backstage ». La demande préparée pour votre outil d'IA ne contient que ce qui concerne cette cible (éléments permis, exemple, structure actuelle), et une réponse qui vise une autre cible est refusée.
- **Onglets contextuels.** Vous choisissez le jeu d'onglets (tableau, graphique, tableau croisé dynamique, image…) dans une liste ; l'assistant y ajoute des onglets, des groupes et des contrôles, sans jamais toucher au reste du ruban.
- **Menu Fichier (Backstage).** L'assistant ajoute des pages, des groupes, des boutons et des commandes rapides, les place avant ou après une commande d'Excel, masque, renomme ou grise une commande d'Excel et pose `onShow` / `onHide`. Seules les 13 commandes du menu Fichier relevées dans Excel sont acceptées. La proposition affiche la liste du menu Fichier telle qu'Excel la montrera.
- Les garde-fous restent les mêmes : aucune connexion, rien n'est appliqué si une opération est refusée, aucun code VBA venant de l'IA, suppressions confirmées, une seule étape « Annuler ».

### Compatibilité

- Les réponses au format de la version 1.8 (sans champ `cible`) restent acceptées pour le ruban.
- Aucun changement des fichiers `.cuis` et `.cuic`.

## 1.9.0

Backstage (menu Fichier) dans la bibliothèque, et aperçu du menu Fichier au plus près d'Excel. Les projets et les classeurs des versions précédentes s'ouvrent et s'enregistrent comme avant.

### Nouveautés

- **Aperçu du menu Fichier proche d'Excel 365.** La liste de gauche reprend celle d'Excel (Accueil, Nouveau, Ouvrir, Informations… Compte, Options) : les commandes masquées disparaissent, les commandes renommées prennent leur nouveau nom, et chaque élément ajouté se place à l'endroit où Excel le met (avant ou après une commande d'Excel, sinon en bas de la liste). La page affiche les groupes, le grand bouton principal, les cases, listes et boutons radio avec les couleurs et la disposition d'Excel. Ordre, couleurs et proportions ont été relevés dans Excel.
- **Backstage dans la bibliothèque.** Le menu Fichier d'un projet s'enregistre comme composant `.cuic`, avec son code VBA, et s'insère dans un autre projet. L'insertion renomme les identifiants et les callbacks avec un préfixe, ne touche jamais les commandes d'Excel, et refuse ce qui entrerait en conflit (deux réglages d'une même commande d'Excel, un second callback `onShow`).
- **Composant fourni « Backstage – Administration ».** Une page « Administration » placée après « Informations » (options d'export, nom du validateur, format de sortie, serveur cible, état du classeur) et une commande rapide « Synchroniser », vérifiées dans Excel avec leur code VBA.
- **Commandes du menu Fichier vérifiées.** Les identifiants des commandes d'Excel du menu Fichier (Accueil `PlaceTabHome`, Informations `TabInfo`, Enregistrer sous `TabSave`, Compte `TabHelp`, Options `ApplicationOptionsDialog`…) portent désormais le badge « vérifié ».

### Corrections

- Un projet sans onglet de ruban (Backstage seul) produisait un XML refusé par Excel ; il est maintenant valide.
- L'aperçu du menu Fichier se réduit pour tenir entièrement dans une fenêtre étroite.

### Simplification

- Le choix de la plateforme visée (« Windows / Autre ») est retiré de la barre de l'aperçu : la compatibilité est évaluée pour Excel sous Windows, la seule plateforme vérifiée. Les projets qui précisaient une autre plateforme s'ouvrent comme avant.

### Compatibilité

- Un fichier `.cuic` de Backstage, ou un projet `.cuis` qui en contient un, ne s'ouvre pas dans les versions antérieures à la 1.9.0 : elles indiquent qu'il a été créé par une version plus récente. Les autres fichiers restent lisibles par la 1.8.

## 1.8.1

Correctif d'affichage. Les projets et les classeurs ne changent pas.

### Corrections

- Quand la fenêtre est étroite, les boutons de l'en-tête et des barres d'outils (module VBA avec plusieurs composants, panneau Propriétés, bibliothèques d'images, catalogue idMso) passent à la ligne au lieu de sortir du cadre ou de se chevaucher. Le bouton « Exporter XML / Enregistrer Office » n'affiche alors que son icône, avec son nom en info-bulle.

## 1.8.0

Mise à jour majeure : projets, composants, modèles, menus clic droit avancés, compléments, traitement par lot et assistant IA. Les classeurs produits avec la 1.0.0 s'ouvrent et s'enregistrent comme avant.

### Nouveautés

- **Projets `.cuis`.** Un ruban s'enregistre indépendamment de tout classeur, se rouvre, puis s'applique au classeur de votre choix (bouton « Appliquer »).
- **Annuler / rétablir.** Chaque modification s'annule (Ctrl+Z) et se rétablit (Ctrl+Y).
- **Onglets contextuels et Backstage.** Les onglets qu'Excel affiche pour un tableau, un graphique ou une image se complètent dans l'éditeur, et le menu Fichier (Backstage) se construit élément par élément.
- **Compatibilité.** Choisissez la version d'Excel visée : chaque commande, icône et menu porte un badge « vérifié », « documenté » ou « inconnu », d'après les listes officielles de Microsoft et des relevés faits dans Excel. Un rapport résume le projet.
- **Composants et bibliothèque.** Un groupe, un onglet ou un menu se range avec son code VBA dans un composant `.cuic`, réutilisable dans tout projet. Des composants sont fournis.
- **Modèles métier.** Gestion de stocks, gestion commerciale et tableau de bord, avec leur ruban, leur code VBA et la macro « Préparer le classeur ».
- **Menus clic droit avancés.** Les 89 menus de la liste officielle de Microsoft ; des entrées affichées seulement là où elles ont un sens (feuille, tableau, colonne, cellule non vide…), avec le code VBA généré ; copie d'une entrée vers d'autres menus ; sous-menus.
- **Compléments `.xlam`.** Ouverture, modification et enregistrement d'un complément ; code VBA généré adapté au complément ; procédure guidée pour le créer dans Excel.
- **Appliquer à plusieurs classeurs** (version de bureau). Contrôle préalable de chaque fichier, copies dans un dossier de sortie, rapport CSV ; les originaux ne sont jamais modifiés.
- **Assistant IA contrôlé.** Décrivez ce que vous voulez : CustomUI Studio prépare une demande pour l'outil d'IA de votre choix, puis vérifie sa réponse avec ses propres règles avant toute modification. L'application ne se connecte à rien et n'envoie rien.

### Corrections

- Un onglet ou un groupe natif d'Excel sans libellé n'est plus renommé à l'export.
- Un retour à la ligne dans une info-bulle est conservé.
- Un attribut qu'un contrôle n'accepte pas est signalé au lieu d'être retiré sans prévenir.
- Un fichier dont l'extension ne correspond pas au contenu (par exemple un `.xlsm` renommé en `.xlam`) est refusé avec une explication, comme le ferait Excel.

## 1.0.0

Première version stable. Les fonctions sont celles de la 0.8.0.

- **Microsoft Store.** Une version installable est en cours de publication sur le Microsoft Store. Elle sera signée par Microsoft et s'installera sans avertissement Windows.
- **Politique de confidentialité.** Elle est publiée dans [CONFIDENTIALITE.md](CONFIDENTIALITE.md) : l'application ne collecte et ne transmet aucune donnée, et fonctionne sans connexion à Internet.
- La version portable reste disponible, gratuite et sans installation.
- Les classeurs produits avec les versions précédentes s'ouvrent et s'enregistrent comme avant.

## 0.8.0

### Nouveautés

- **Interface en français ou en anglais.** Au premier lancement, l'application demande la langue ; une liste en haut de la fenêtre permet d'en changer à tout moment. Le choix est mémorisé.
- **Propriétés pilotées par VBA.** Dans les propriétés d'un contrôle, cochez ce qui doit pouvoir changer pendant l'utilisation du classeur : libellé, activé / grisé, visible / masqué, info-bulles, image. Dans vos macros, une seule ligne suffit ensuite :

  ```vb
  SetControlState "btnExporter", "enabled", False
  ```

  Le module `RibbonDESIGNER` généré mémorise la valeur et actualise le ruban. Tant que vous n'appelez pas `SetControlState`, la valeur saisie dans l'éditeur s'applique.
- **Glisser-déposer dans la Structure du ruban.** Onglets, groupes et contrôles se réorganisent à la souris : avant, après ou dans un autre élément. L'application refuse ce qu'Excel n'accepterait pas et dit pourquoi. Les flèches monter / descendre restent disponibles.
- **Menus clic droit : masquer davantage de commandes d'Excel.** Certaines commandes (Collage spécial, Insérer, Effacer le contenu, Filtrer, Trier…) ne peuvent pas être masquées par le ruban lui-même. Elles portent l'étiquette « VBA » : en les décochant, le module `RibbonDESIGNER` reçoit les procédures `MasquerCommandesNatives` et `RestaurerCommandesNatives`. Ce masquage agit sur tout Excel tant que le classeur est actif ; trois courtes procédures à coller dans `ThisWorkbook`, fournies en tête du module, rendent les commandes aux autres classeurs.
- **Code VBA plus parlant.** Les procédures générées contiennent des exemples en commentaire pour les cas courants.
- **Aperçu du clic droit.** Une croix permet de le fermer pour revoir le ruban.

### Corrections

- **Image supprimée de la bibliothèque :** elle est retirée du classeur à l'enregistrement, avec sa liaison. Auparavant elle restait dans l'archive.
- **Image réimportée sous le même nom :** le classeur reçoit bien la nouvelle image.
- **Image utilisée par un contrôle :** la supprimer demande confirmation en nommant les contrôles concernés.
- **Changement de format (CustomUI 2006 vers 2009, ou l'inverse) :** l'ancien fichier du ruban ne reste plus dans le classeur.
- **Panneau « Problèmes » :** textes clarifiés.

Les propriétés pilotées par VBA, le masquage par VBA et les classeurs dont on a supprimé ou remplacé des images ont été essayés dans Excel pour Microsoft 365 (version 16.0), en français.

## 0.6.0

### Nouveautés

- **Vérification du XML avec les schémas de Microsoft.** L'éditeur XML contrôle votre ruban en continu, hors ligne, avec les schémas CustomUI 2006 et 2009.
  - Panneau « Problèmes » : chaque défaut a sa ligne, une explication et, quand c'est possible, la correction à faire. Un clic place le curseur dessus.
  - Les erreurs sont soulignées dans l'éditeur.
  - L'enregistrement dans un classeur et l'export XML sont bloqués tant qu'il reste une erreur, pour ne plus obtenir un ruban qui disparaît sans explication dans Excel.
- **Menus clic droit : masquer les commandes d'Excel en un clic**, pour la cellule, l'en-tête de ligne, l'en-tête de colonne et le tableau.
  - Liste des commandes d'Excel avec une case à cocher chacune.
  - Préréglages : « Menu minimal », « Sans copier-coller », « Sans insertion ni suppression », « Tout afficher ».
  - L'aperçu montre le menu tel qu'Excel l'affichera. Cliquez sur une commande d'Excel pour la masquer ou placer un bouton juste avant ou après.

### Corrections

- **Code VBA généré.** Ces défauts empêchaient le ruban de s'afficher ou provoquaient « Incompatibilité de type » dans Excel :
  - les procédures `get…` (getLabel, getEnabled, getVisible…) et celle d'une commande d'Excel interceptée déclaraient un type pour leur valeur de retour ;
  - `getVisible` et `getEnabled` renvoyaient Faux par défaut : le contrôle disparaissait ou était grisé ;
  - `loadImage` était généré comme une fonction ;
  - `onAction` d'une liste déroulante ou d'une galerie ne recevait pas l'index de l'élément choisi ;
  - `onChange` d'une zone de texte nommait mal son paramètre.
- **Boutons bascule et cases à cocher :** chacun garde son propre état (auparavant, tous partageaient le même).
- **Attributs refusés par Excel :** l'éditeur ne propose plus 26 attributs que le format interdit (par exemple une image sur une case à cocher, un libellé sur un bouton fractionné). Excel refusait alors tout le ruban.
- **Libellé fixe et callback ensemble** (`label` et `getLabel`, `visible` et `getVisible`…) : seul le callback est écrit, et l'éditeur le signale.
- **Aperçu du clic droit :** les images personnalisées s'affichent.
- **Identifiant d'un groupe ou d'un onglet :** le panneau Propriétés ne perd plus la sélection pendant la saisie.
- **Images chargées par VBA :** un ruban qui utilise `loadImage` peut être enregistré sans embarquer les images.

Le code VBA généré, le refus des attributs interdits et les menus clic droit ont été essayés dans Excel pour Microsoft 365 (version 16.0), en français.

## 0.4.0

### Nouveautés

- **Menus contextuels (clic droit) :** nouvelle section « Menus contextuels » dans la Structure du ruban.
  - Choix du menu d'Excel à compléter : cellule, en-tête de ligne, en-tête de colonne, tableau structuré, tableau croisé dynamique, forme (ou tout autre nom).
  - Ajout de boutons, cases à cocher, sous-menus, séparateurs… : seuls les contrôles autorisés dans un menu clic droit sont proposés.
  - Position réglable : avant ou après une commande d'Excel (Couper, Copier, Coller…), ou à la fin du menu.
  - Aperçu du menu clic droit, callbacks VBA générés, relecture depuis l'éditeur XML.
  - Format CustomUI 2009 (Office 2010 et suivants) : l'application propose de basculer si votre projet est en 2006.

### Corrections

- **Aucune perte de données à la synchronisation XML :** les menus contextuels, le Backstage et les onglets contextuels présents dans un classeur étaient supprimés en cliquant sur « Synchroniser ». Ils sont conservés et réécrits à la place exigée par Office.
- **Changement d'identifiant d'un contrôle :** le panneau Propriétés ne perd plus le contrôle sélectionné pendant la saisie de son identifiant.

## 0.3.0

Première version publique.

- Construction visuelle du ruban : seuls les éléments autorisés par Excel sont proposés.
- Formats CustomUI 2006 (Office 2007) et 2009 (Office 2010 et suivants).
- Aperçu en direct avec listes, menus et info-bulles.
- Bibliothèque imageMso, images personnalisées et catalogue des commandes natives d'Excel.
- Module VBA `RibbonDESIGNER` généré, importé par la macro `ImporterModuleRibbonDESIGNER`.
- Éditeur XML synchronisé, contrôle complet avant l'enregistrement.
- Classeur non reconstruit : seul le ruban change, vos feuilles, données et macros VBA restent intactes.
