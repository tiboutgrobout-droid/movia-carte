/*
 * MOVIA — Carte des interventions — réglages.
 * Fond de carte : OpenStreetMap pendant le pilote (développement et faible usage, attribution obligatoire).
 * Avant la commercialisation : remplacer par un fournisseur sous contrat (MapTiler, IGN Géoplateforme…) en changeant
 * seulement les deux lignes ci-dessous. Rien d'autre n'est à modifier.
 */
window.MOVIA_CONFIG = {
  fondUrl: "https://tile.openstreetmap.org/{z}/{x}/{y}.png",
  fondAttribution: '&copy; <a href="https://www.openstreetmap.org/copyright" target="_blank" rel="noopener">contributeurs OpenStreetMap</a>',
  zoomMax: 18,
  rayonRegroupementPx: 28,          // repères plus proches que cela à l'écran : regroupés (cliquer pour zoomer)
  tableCarte: "Carte",              // écrite par le script MOVIA (module M26)
  tableAgences: "Agences",
  celluleAgence: { feuille: "Accueil 2PA", adresse: "B15" }   // agence choisie dans l'Accueil (règle : une agence à la fois)
};
