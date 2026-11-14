# Historique des versions

La dernière version se télécharge sur la page [Releases](https://github.com/CustomUI-Studio/CustomUI-Studio/releases/latest).

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
