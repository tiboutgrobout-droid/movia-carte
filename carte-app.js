/*
 * MOVIA — Carte des interventions — volet Excel (lecture seule).
 * Lit la table « Carte » (écrite par le script MOVIA, module M26), la table Agences et l'agence choisie dans l'Accueil,
 * puis affiche la journée sur une carte Leaflet. AUCUNE écriture dans le classeur (manifeste : ReadDocument).
 * Ne calcule ni itinéraire, ni trajet, ni horaire : consultation uniquement.
 */
(function () {
  "use strict";
  var C = window.MOVIA_CONFIG, M = window.MoviaCarte;
  var etat = { alerteCarte: "", lignes: [], agences: [], agenceAccueil: "", agence: "", date: 0, visibles: {}, lieux: [], carte: null, couche: null, coucheAgence: null, minuterie: null };
  var $ = function (id) { return document.getElementById(id); };
  var COULEUR = /^#[0-9A-Fa-f]{6}$/;
  function couleurSure(c, defaut) { return COULEUR.test(String(c || "")) ? c : (defaut || "#1F4A3D"); }

  function message(texte, info) {
    var m = $("message");
    if (!texte && etat.alerteCarte) { texte = etat.alerteCarte; info = false; }   // la carte indisponible reste signalée
    if (!texte) { m.hidden = true; m.textContent = ""; return; }
    m.hidden = false; m.textContent = texte; m.className = "message" + (info ? " info" : "");
  }

  // ---------------------------------------------------------------- Lecture du classeur (lecture seule)
  function lireClasseur() {
    return Excel.run(function (ctx) {
      var t = ctx.workbook.tables.getItemOrNullObject(C.tableCarte);
      var ta = ctx.workbook.tables.getItemOrNullObject(C.tableAgences);
      var f = ctx.workbook.worksheets.getItemOrNullObject(C.celluleAgence.feuille);
      t.load("isNullObject"); ta.load("isNullObject"); f.load("isNullObject");
      return ctx.sync().then(function () {
        var r = null, ra = null, ca = null;
        if (!t.isNullObject) { r = t.getRange(); r.load("values"); }
        if (!ta.isNullObject) { ra = ta.getRange(); ra.load("values"); }
        if (!f.isNullObject) { ca = f.getRange(C.celluleAgence.adresse); ca.load("values"); }
        return ctx.sync().then(function () {
          return { carte: r ? r.values : null, agences: ra ? ra.values : [], agenceAccueil: ca ? String(ca.values[0][0] || "").trim() : "" };
        });
      });
    });
  }

  /** Actualisation automatique quand le script réécrit la table Carte (si la version d'Excel le permet). */
  function surveiller() {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) return;
    Excel.run(function (ctx) {
      var t = ctx.workbook.tables.getItemOrNullObject(C.tableCarte);
      t.load("isNullObject");
      return ctx.sync().then(function () {
        if (t.isNullObject) return null;
        t.onChanged.add(function () { clearTimeout(etat.minuterie); etat.minuterie = setTimeout(function () { charger(false); }, 1500); return Promise.resolve(); });
        return ctx.sync();
      });
    }).catch(function () { /* sans surveillance : bouton Actualiser */ });
  }

  function charger(premier) {
    return lireClasseur().then(function (d) {
      if (!d.carte) {
        etat.lignes = [];
        message("Table « Carte » introuvable : lancez une fois « Générer le planning » avec le script MOVIA V2.5-O5.1 (ou plus récent), puis « Actualiser ».");
      } else {
        var manque = M.colonnesManquantes(d.carte);
        etat.lignes = manque.length ? [] : M.lireTable(d.carte);
        message(manque.length ? "Table « Carte » incomplète (colonnes manquantes : " + manque.join(", ") + ") : mettez à jour les deux scripts MOVIA." : "");
      }
      etat.agences = M.agences(etat.lignes, M.lireTable(d.agences));
      etat.agenceAccueil = d.agenceAccueil;
      if (premier || !etat.agence) {
        etat.agence = d.agenceAccueil || (etat.agences[0] ? etat.agences[0].id : "");
        etat.date = M.jourParDefaut(etat.lignes, etat.agence, M.aujourdhuiSerie());
      }
      remplirAgences();
      var h = new Date();
      $("version").textContent = " · Données lues à " + h.getHours() + "h" + (h.getMinutes() < 10 ? "0" : "") + h.getMinutes() + " · volet " + M.VERSION;
      afficher(premier);
    }).catch(function (e) {
      message("Lecture du classeur impossible : " + (e && e.message ? e.message : e) + ". Réessayez avec « Actualiser ».");
    });
  }

  // ---------------------------------------------------------------- Carte
  function initCarte() {
    if (!window.L) { etat.alerteCarte = "La carte n'a pas pu être chargée (bibliothèque Leaflet inaccessible) : vérifiez la connexion Internet, puis rouvrez le volet. La liste des équipes reste affichée."; message(""); return false; }
    etat.carte = L.map("carte", { zoomControl: true }).setView([46.6, 2.4], 5);
    L.tileLayer(C.fondUrl, { maxZoom: C.zoomMax, attribution: C.fondAttribution }).addTo(etat.carte);
    etat.coucheAgence = L.layerGroup().addTo(etat.carte);
    etat.couche = L.layerGroup().addTo(etat.carte);
    etat.carte.on("zoomend", dessinerReperes);
    return true;
  }

  function iconeRepere(lieu) {
    var c1 = couleurSure(lieu.couleur), n = lieu.interventions.length;
    var autres = lieu.equipes.length > 1;   // plusieurs équipes au même endroit : centre ivoire cerclé de terracotta
    var svg = '<svg width="26" height="36" viewBox="0 0 26 36" aria-hidden="true">'
      + '<path d="M13 35C13 35 2 21.5 2 13a11 11 0 0 1 22 0c0 8.5-11 22-11 22z" fill="' + c1 + '" stroke="#F7F3EA" stroke-width="2"/>'
      + '<circle cx="13" cy="13" r="4.6" fill="#F7F3EA"' + (autres ? ' stroke="#C8734F" stroke-width="2.2"' : "") + "/></svg>";
    return L.divIcon({ className: "", html: '<div class="repere">' + svg + (n > 1 ? '<span class="badge">' + n + "</span>" : "") + "</div>",
      iconSize: [26, 36], iconAnchor: [13, 35], popupAnchor: [0, -32] });
  }

  function dessinerReperes() {
    if (!etat.carte) return;
    etat.couche.clearLayers();
    var pts = etat.lieux.map(function (l) { return etat.carte.latLngToLayerPoint([l.lat, l.lon]); });
    var groupes = M.regrouperProches(pts, C.rayonRegroupementPx);
    groupes.forEach(function (g) {
      if (g.length === 1) {
        var lieu = etat.lieux[g[0]];
        var titre = lieu.interventions.map(function (i) { return i.IdIntervention; }).join(", ");
        L.marker([lieu.lat, lieu.lon], { icon: iconeRepere(lieu), title: titre, keyboard: true })
          .bindPopup(M.ficheHtml(lieu), { maxWidth: 280, maxHeight: 320, autoPanPadding: [20, 20] }).addTo(etat.couche);
        return;
      }
      var membres = g.map(function (i) { return etat.lieux[i]; });
      var nb = membres.reduce(function (t, l) { return t + l.interventions.length; }, 0);
      var lat = membres.reduce(function (t, l) { return t + l.lat; }, 0) / membres.length;
      var lon = membres.reduce(function (t, l) { return t + l.lon; }, 0) / membres.length;
      var bornes = L.latLngBounds(membres.map(function (l) { return [l.lat, l.lon]; }));
      var m = L.marker([lat, lon], { icon: L.divIcon({ className: "", html: '<div class="groupe" title="' + nb + ' interventions proches : cliquer pour zoomer">' + nb + "</div>", iconSize: [34, 34], iconAnchor: [17, 17] }), title: nb + " interventions proches", keyboard: true });
      m.on("click", function () {
        if (etat.carte.getZoom() >= C.zoomMax - 1) {
          m.bindPopup(membres.map(M.ficheHtml).join(""), { maxWidth: 280, maxHeight: 320 }).openPopup();
        } else etat.carte.fitBounds(bornes, { padding: [40, 40], maxZoom: C.zoomMax });
      });
      m.addTo(etat.couche);
    });
  }

  function agenceCourante() { return etat.agences.filter(function (a) { return a.id === etat.agence; })[0]; }

  function recentrer() {
    if (!etat.carte) return;
    var ag = agenceCourante();
    if (etat.lieux.length === 0) {
      if (ag && ag.lat !== undefined && ag.lon !== undefined) etat.carte.setView([ag.lat, ag.lon], 8); else etat.carte.setView([46.6, 2.4], 5);
      return;
    }
    if (etat.lieux.length === 1) { etat.carte.setView([etat.lieux[0].lat, etat.lieux[0].lon], 11); return; }
    etat.carte.fitBounds(L.latLngBounds(etat.lieux.map(function (l) { return [l.lat, l.lon]; })), { padding: [30, 30], maxZoom: 12 });
  }

  // ---------------------------------------------------------------- Affichage
  function remplirAgences() {
    var s = $("agence"); s.innerHTML = "";
    etat.agences.forEach(function (a) {
      var o = document.createElement("option");
      o.value = a.id; o.textContent = a.id + " — " + a.nom + (a.id === etat.agenceAccueil ? " (Accueil)" : "");
      s.appendChild(o);
    });
    s.value = etat.agence;
  }

  function afficher(recentrage) {
    var j = M.donneesJour(etat.lignes, etat.agence, etat.date);
    j.equipes.forEach(function (e) { if (etat.visibles[e.numero] === undefined) etat.visibles[e.numero] = true; });
    etat.lieux = M.lieux(j, etat.visibles);
    $("date").value = M.serieVersIso(etat.date);
    $("jourLibelle").textContent = M.jourLong(etat.date) + (etat.date === M.aujourdhuiSerie() ? " · aujourd'hui" : "");
    var visibles = etat.lieux.reduce(function (t, l) { return t + l.interventions.length; }, 0);
    $("total").textContent = visibles;
    $("totalLibelle").textContent = (visibles > 1 ? "interventions" : "intervention") + " sur la carte"
      + (visibles !== j.carte.length ? " (sur " + j.carte.length + " ; équipes masquées)" : "");

    var leg = $("legende"); leg.innerHTML = "";
    if (j.equipes.length === 0) { var vide = document.createElement("li"); vide.className = "agents"; vide.textContent = "Aucune équipe planifiée ce jour."; leg.appendChild(vide); }
    j.equipes.forEach(function (e) {
      var li = document.createElement("li"); if (!etat.visibles[e.numero]) li.className = "masquee";
      li.innerHTML = '<label><input type="checkbox"' + (etat.visibles[e.numero] ? " checked" : "") + ' data-equipe="' + e.numero + '">'
        + '<span class="pastille" style="background:' + couleurSure(e.couleur) + '"></span>'
        + '<span><span class="nom">' + M.echapper(e.libelle) + "</span> · " + e.interventions + " interv." + (e.horsCarte ? " (+" + e.horsCarte + " hors carte)" : "")
        + '<br><span class="agents">' + M.echapper(e.agents) + "</span></span></label>";
      leg.appendChild(li);
    });

    var hors = M.horsCarte(j);
    $("horsBloc").hidden = hors.length === 0;
    $("horsNb").textContent = hors.length;
    var ul = $("hors"); ul.innerHTML = "";
    hors.forEach(function (h) {
      var li = document.createElement("li");
      li.innerHTML = '<span class="ref">' + M.echapper(h.reference) + "</span> · " + M.echapper(h.client) + (h.equipe ? " · " + M.echapper(h.equipe) : "")
        + '<br><span class="detail">' + M.echapper(h.detail) + "</span>";
      ul.appendChild(li);
    });

    if (etat.agenceAccueil && etat.agence !== etat.agenceAccueil) message("Agence de l'Accueil : " + etat.agenceAccueil + ". Vous consultez le dernier planning publié de " + etat.agence + ".", true);
    else if ($("message").className.indexOf("info") >= 0) message("");

    if (etat.coucheAgence && window.L) {
      etat.coucheAgence.clearLayers();
      var ag = agenceCourante();
      if (ag && ag.lat !== undefined && ag.lon !== undefined) {
        L.marker([ag.lat, ag.lon], { icon: L.divIcon({ className: "", html: '<div class="agence-repere">A</div>', iconSize: [20, 20], iconAnchor: [10, 10] }), title: "Agence " + ag.nom, keyboard: false })
          .bindTooltip("Agence " + M.echapper(ag.nom)).addTo(etat.coucheAgence);
      }
    }
    dessinerReperes();
    if (recentrage) recentrer();
  }

  function changerJour(d, recentrage) { if (d === undefined) return; etat.date = d; afficher(recentrage !== false); }

  function brancher() {
    $("agence").addEventListener("change", function (e) {
      etat.agence = e.target.value; etat.visibles = {};
      etat.date = M.jourParDefaut(etat.lignes, etat.agence, M.aujourdhuiSerie()); afficher(true);
    });
    $("date").addEventListener("change", function (e) { var d = M.isoVersSerie(e.target.value); if (d !== undefined) { etat.visibles = {}; changerJour(d); } });
    $("jourPrec").addEventListener("click", function () { etat.visibles = {}; changerJour(etat.date - 1); });
    $("jourSuiv").addEventListener("click", function () { etat.visibles = {}; changerJour(etat.date + 1); });
    $("jourPlanPrec").addEventListener("click", function () {
      var d = M.jourVoisinAvecInterventions(etat.lignes, etat.agence, etat.date, -1);
      if (d === undefined) message("Aucune journée planifiée avant le " + M.dateCourte(etat.date) + ".", true); else { etat.visibles = {}; changerJour(d); }
    });
    $("jourPlanSuiv").addEventListener("click", function () {
      var d = M.jourVoisinAvecInterventions(etat.lignes, etat.agence, etat.date, 1);
      if (d === undefined) message("Aucune journée planifiée après le " + M.dateCourte(etat.date) + ".", true); else { etat.visibles = {}; changerJour(d); }
    });
    $("actualiser").addEventListener("click", function () { charger(false); });
    $("recentrer").addEventListener("click", recentrer);
    $("toutes").addEventListener("click", function () { Object.keys(etat.visibles).forEach(function (k) { etat.visibles[k] = true; }); afficher(false); });
    $("legende").addEventListener("change", function (e) {
      var n = e.target && e.target.getAttribute("data-equipe");
      if (n === null) return;
      etat.visibles[n] = e.target.checked; afficher(false);
    });
  }

  window.MoviaVolet = { etat: etat };   // accès de diagnostic (tests), lecture seule

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Excel) { message("Ce volet fonctionne dans Excel."); return; }
    brancher();
    initCarte();
    charger(true).then(surveiller);
    window.addEventListener("resize", function () { if (etat.carte) etat.carte.invalidateSize(); });
  });
})();
