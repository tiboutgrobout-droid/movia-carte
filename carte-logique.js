/*
 * MOVIA — Carte des interventions — logique pure (aucun accès à Excel ni à la carte).
 * Lit le contenu de la table « Carte » écrite par le script MOVIA (module M26) et prépare ce que le volet affiche :
 * journée, interventions par lieu, équipes, légende, interventions hors carte, fiche d'intervention.
 * Ce fichier sera réutilisé tel quel dans MOVIA Windows. Version 1.1.0 (O5.2 : couleur par défaut de la charte V3).
 */
(function (racine, fabrique) {
  if (typeof module === "object" && module.exports) module.exports = fabrique();
  else racine.MoviaCarte = fabrique();
}(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var VERSION = "1.1.0";
  var JOURS = ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"];
  var ORIGINE = Date.UTC(1899, 11, 30);   // numéro de série Excel 0
  var JOUR_MS = 86400000;
  var COLONNES_REQUISES = ["AgencyID", "Date", "Categorie", "IdIntervention", "Latitude", "Longitude", "NumeroEquipe", "Equipe", "CouleurEquipe", "Agents"];

  function texte(v) { return v === undefined || v === null ? "" : String(v).trim(); }
  function nombre(v) { if (typeof v === "number") return v; var t = texte(v).replace(",", "."); if (t === "") return undefined; var n = Number(t); return isFinite(n) ? n : undefined; }
  function deux(n) { return (n < 10 ? "0" : "") + n; }

  /** Numéro de série Excel → { a, m, j, js } (date civile, sans fuseau). */
  function serieVersDate(s) { var d = new Date(ORIGINE + Math.floor(s) * JOUR_MS); return { a: d.getUTCFullYear(), m: d.getUTCMonth() + 1, j: d.getUTCDate(), js: d.getUTCDay() }; }
  function dateVersSerie(a, m, j) { return Math.round((Date.UTC(a, m - 1, j) - ORIGINE) / JOUR_MS); }
  /** « 2026-11-19 » (champ date du navigateur) ↔ numéro de série. */
  function isoVersSerie(iso) { var x = /^(\d{4})-(\d{2})-(\d{2})$/.exec(texte(iso)); return x ? dateVersSerie(+x[1], +x[2], +x[3]) : undefined; }
  function serieVersIso(s) { var d = serieVersDate(s); return d.a + "-" + deux(d.m) + "-" + deux(d.j); }
  function dateCourte(s) { var d = serieVersDate(s); return deux(d.j) + "/" + deux(d.m) + "/" + d.a; }
  function jourLong(s) { return JOURS[serieVersDate(s).js] + " " + dateCourte(s); }
  function aujourdhuiSerie(maintenant) { var d = maintenant || new Date(); return dateVersSerie(d.getFullYear(), d.getMonth() + 1, d.getDate()); }

  /** Valeurs d'une table Excel (ligne d'en-têtes + lignes) → objets ; lignes vides ignorées. */
  function lireTable(valeurs) {
    if (!valeurs || valeurs.length === 0) return [];
    var e = valeurs[0].map(texte), res = [];
    for (var i = 1; i < valeurs.length; i++) {
      var l = valeurs[i];
      if (l.every(function (v) { return texte(v) === ""; })) continue;
      var o = {};
      for (var k = 0; k < e.length; k++) o[e[k]] = l[k];
      res.push(o);
    }
    return res;
  }
  function colonnesManquantes(valeurs) {
    var e = valeurs && valeurs.length ? valeurs[0].map(texte) : [];
    return COLONNES_REQUISES.filter(function (c) { return e.indexOf(c) < 0; });
  }

  /** Liste des agences proposées : celles de la table Agences, puis celles de la carte, sans doublon. */
  function agences(lignes, tableAgences) {
    var res = [], vus = {};
    (tableAgences || []).forEach(function (a) { var id = texte(a.AgencyID); if (id && !vus[id]) { vus[id] = 1; res.push({ id: id, nom: texte(a.NomAgence) || id, lat: nombre(a.Latitude), lon: nombre(a.Longitude) }); } });
    (lignes || []).forEach(function (l) { var id = texte(l.AgencyID); if (id && !vus[id]) { vus[id] = 1; res.push({ id: id, nom: id }); } });
    return res;
  }

  function deLAgence(lignes, agence) { return (lignes || []).filter(function (l) { return texte(l.AgencyID) === agence; }); }
  function estSurCarte(l) { return texte(l.Categorie) === "CARTE"; }

  /** Journée ouverte par défaut : aujourd'hui s'il y a des interventions sur la carte, sinon la prochaine (même règle que le Tableau de bord). */
  function jourParDefaut(lignes, agence, aujourdhui) {
    var dates = deLAgence(lignes, agence).filter(estSurCarte).map(function (l) { return nombre(l.Date); }).filter(function (d) { return d !== undefined && d >= aujourdhui; });
    return dates.length ? Math.min.apply(null, dates) : aujourdhui;
  }
  /** Journée précédente / suivante ayant au moins une intervention (sur la carte ou hors carte) ; undefined s'il n'y en a pas. */
  function jourVoisinAvecInterventions(lignes, agence, date, sens) {
    var dates = deLAgence(lignes, agence).map(function (l) { return nombre(l.Date); }).filter(function (d) { return d !== undefined && (sens > 0 ? d > date : d < date); });
    if (!dates.length) return undefined;
    return sens > 0 ? Math.min.apply(null, dates) : Math.max.apply(null, dates);
  }

  /** Tout ce qu'affiche le volet pour une agence et une journée. */
  function donneesJour(lignes, agence, date) {
    var jour = deLAgence(lignes, agence).filter(function (l) { return nombre(l.Date) === date; });
    var carte = jour.filter(estSurCarte), hors = jour.filter(function (l) { return !estSurCarte(l); });
    var equipes = {}, liste = [];
    carte.concat(hors).forEach(function (l) {
      var n = nombre(l.NumeroEquipe);
      if (n === undefined) return;
      if (!equipes[n]) { equipes[n] = { numero: n, libelle: texte(l.Equipe) || ("Équipe " + n), couleur: texte(l.CouleurEquipe) || "#3B4A6B", agents: texte(l.Agents), interventions: 0, horsCarte: 0 }; liste.push(equipes[n]); }
      if (estSurCarte(l)) equipes[n].interventions++; else equipes[n].horsCarte++;
    });
    liste.sort(function (a, b) { return a.numero - b.numero; });
    return { agence: agence, date: date, carte: carte, hors: hors, equipes: liste };
  }

  /** Interventions sur la carte regroupées par lieu (même adresse = mêmes coordonnées), limitées aux équipes affichées. */
  function lieux(jour, equipesVisibles) {
    var res = [], parCle = {};
    jour.carte.forEach(function (l) {
      var n = nombre(l.NumeroEquipe);
      if (equipesVisibles && n !== undefined && !equipesVisibles[n]) return;
      var lat = nombre(l.Latitude), lon = nombre(l.Longitude);
      if (lat === undefined || lon === undefined) return;   // sécurité : jamais de position inventée
      var cle = texte(l.CleLieu) || (lat.toFixed(5) + ";" + lon.toFixed(5));
      if (!parCle[cle]) { parCle[cle] = { cle: cle, lat: lat, lon: lon, interventions: [], equipes: [] }; res.push(parCle[cle]); }
      parCle[cle].interventions.push(l);
      if (n !== undefined && parCle[cle].equipes.indexOf(n) < 0) parCle[cle].equipes.push(n);
    });
    res.forEach(function (x) { x.equipes.sort(function (a, b) { return a - b; }); x.couleur = texte(x.interventions[0].CouleurEquipe) || "#3B4A6B"; });
    return res;
  }

  /**
   * Regroupe les repères trop proches à l'écran (distance en pixels < rayon) : renvoie des groupes d'indices.
   * Fonction pure : la carte lui donne les positions en pixels au zoom courant.
   */
  function regrouperProches(points, rayon) {
    var groupes = [], pris = [];
    for (var i = 0; i < points.length; i++) {
      if (pris[i]) continue;
      var g = [i]; pris[i] = true;
      for (var j = i + 1; j < points.length; j++) {
        if (pris[j]) continue;
        for (var k = 0; k < g.length; k++) {
          var dx = points[g[k]].x - points[j].x, dy = points[g[k]].y - points[j].y;
          if (dx * dx + dy * dy < rayon * rayon) { g.push(j); pris[j] = true; break; }
        }
      }
      groupes.push(g);
    }
    return groupes;
  }

  function echapper(v) {
    return texte(v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }
  var GEO = { SAISIE: "", VALIDE: "", A_VERIFIER: "Position à vérifier", APPROXIMATIF: "Position approximative" };

  /** Fiche (HTML échappé) d'un lieu : une rubrique par intervention. */
  function ficheHtml(lieu) {
    var l0 = lieu.interventions[0];
    var h = '<div class="fiche"><div class="fiche-client">' + echapper(l0.Client || l0.ClientID) + "</div>"
      + '<div class="fiche-adresse">' + echapper(l0.Adresse || "Adresse non renseignée") + "</div>";
    var geo = texte(l0.StatutGeo).toUpperCase();
    if (GEO[geo] === undefined && geo !== "") h += '<div class="fiche-alerte">Position : ' + echapper(geo) + "</div>";
    else if (GEO[geo]) h += '<div class="fiche-alerte">' + GEO[geo] + "</div>";
    if (lieu.interventions.length > 1) h += '<div class="fiche-nb">' + lieu.interventions.length + " interventions à cette adresse</div>";
    lieu.interventions.forEach(function (l) {
      var chantier = texte(l.JourChantier) ? " · chantier " + echapper(l.JourChantier) : "";
      var effectif = texte(l.Effectif) + (texte(l.EffectifRequis) && texte(l.EffectifRequis) !== texte(l.Effectif) ? " (requis : " + echapper(l.EffectifRequis) + ")" : "");
      h += '<div class="fiche-int">'
        + '<div class="fiche-ref"><span class="pastille" style="background:' + echapper(l.CouleurEquipe || "#3B4A6B") + '"></span>'
        + echapper(l.IdIntervention) + " — " + echapper(l.Prestation || l.PrestationID) + "</div>"
        + "<table>"
        + "<tr><th>Date</th><td>" + echapper(jourLong(nombre(l.Date))) + chantier + "</td></tr>"
        + "<tr><th>Horaires</th><td>" + echapper(l.Debut) + " – " + echapper(l.Fin) + "</td></tr>"
        + "<tr><th>Équipe</th><td>" + echapper(l.Equipe) + "</td></tr>"
        + "<tr><th>Agents</th><td>" + echapper(l.Agents) + "</td></tr>"
        + "<tr><th>Effectif</th><td>" + effectif + "</td></tr>"
        + "<tr><th>Statut</th><td>" + echapper(l.Statut) + "</td></tr>"
        + "</table></div>";
    });
    return h + "</div>";
  }

  /** Lignes de la rubrique « Hors carte ». */
  function horsCarte(jour) {
    return jour.hors.map(function (l) {
      return { reference: texte(l.IdIntervention), client: texte(l.Client || l.ClientID), motif: texte(l.Motif), detail: texte(l.Detail), equipe: texte(l.Equipe), couleur: texte(l.CouleurEquipe) };
    });
  }

  return {
    VERSION: VERSION, lireTable: lireTable, colonnesManquantes: colonnesManquantes, agences: agences, jourParDefaut: jourParDefaut,
    jourVoisinAvecInterventions: jourVoisinAvecInterventions, donneesJour: donneesJour, lieux: lieux, regrouperProches: regrouperProches,
    ficheHtml: ficheHtml, horsCarte: horsCarte, echapper: echapper, serieVersDate: serieVersDate, dateVersSerie: dateVersSerie,
    isoVersSerie: isoVersSerie, serieVersIso: serieVersIso, dateCourte: dateCourte, jourLong: jourLong, aujourdhuiSerie: aujourdhuiSerie
  };
}));
