# MOVIA — Carte des interventions (volet Excel) — version 1.0.0

Volet qui s'ouvre à droite du classeur et affiche, sur une carte, les interventions du **planning publié** d'une journée.
Il lit la table « Carte » écrite par les scripts MOVIA V2.5-O5.1 (ou plus récents). **Il ne modifie jamais le classeur.**

## Contenu

| Fichier | Rôle |
|---|---|
| taskpane.html, carte.css | Page du volet (identité MOVIA) |
| carte-logique.js | Calculs purs : journée, équipes, lieux, regroupements, fiche (réutilisable dans MOVIA Windows) |
| carte-app.js | Lecture du classeur (lecture seule) et affichage Leaflet |
| config.js | Réglages, dont le **fond de carte** (à changer avant la commercialisation) |
| commands.html, aide.html, index.html, icon-*.png | Pages techniques, aide, icônes (tous les fichiers au même niveau, sans dossier) |
| manifest.xml | Fichier d'installation dans Excel (n'a pas besoin d'être sur GitHub) |

## 1. Mettre les fichiers en ligne (GitHub Pages, gratuit) — une seule fois

1. Sur github.com : **New repository** → nom `movia-carte` → **Public** → *Create repository*.
   Le dépôt ne contient que le programme ; aucune donnée du classeur n'y est envoyée.
2. Dans le dépôt : **Add file → Upload files**, lien **choose your files** → sélectionner **les 14 fichiers** (Ctrl + A) → *Ouvrir* → *Commit changes*. Il n'y a aucun dossier à envoyer.
3. **Settings → Pages** → *Source* : « Deploy from a branch », branche `main`, dossier `/ (root)` → *Save*.
4. Attendre 1 à 2 minutes, puis ouvrir `https://VOTRE-IDENTIFIANT.github.io/movia-carte/aide.html` : la page d'aide MOVIA doit s'afficher.

## 2. Installer le volet dans Excel

Le manifeste doit contenir votre identifiant GitHub à la place de `VOTRE-IDENTIFIANT`.

- **Excel pour le web** (recommandé pour l'installation) : ouvrir le classeur → onglet **Accueil** → bouton **Compléments** (à droite du ruban) → **Autres compléments** → **Mes compléments** → lien **Charger mon complément** (en haut à droite) → choisir `manifest.xml` → *Charger*.
- **Excel sur Windows** : même chemin (**Accueil → Compléments → Autres compléments → Mes compléments → Gérer mes compléments → Charger un complément**), s'il est proposé.

Ensuite : onglet **Accueil**, groupe **MOVIA**, bouton **Carte des interventions**.
Si le bouton disparaît (cache du navigateur vidé), recharger le manifeste de la même façon.

## 3. Mettre à jour le volet

Remplacer les fichiers dans le dépôt GitHub (Upload files) : le volet est à jour à la prochaine ouverture. Le manifeste ne change pas.

## Tests (dossier tests/, hors dépôt)

- `node tests/tests_logique.js` : 23 vérifications sur la table Carte réelle produite par le script (simulateur).
- `node tests/tests_volet.js` : 25 vérifications dans Chrome, avec un faux Excel (aucune écriture constatée) et une carte simplifiée.
- Non vérifiable hors Excel : le fond de carte réel et le comportement exact d'Excel (installation, actualisation automatique).
