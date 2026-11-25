/**
 * Climat — Amaltis · serveur (Google Apps Script)
 * Application personnelle de Phil : cibles du jour, consignes réellement entrées
 * dans l'automate, rappels de changement, relevés manuels, comparaison
 * cible / réglé / mesuré et registre GACP exportable.
 *
 * Lecture seule : registres de culture, Plan_VPD (dernière version), calendrier.
 * Écriture : la feuille « Climat — Amaltis » (créée par setup) et, sur demande,
 * le statut « fait » d'une tâche Climat du calendrier (+ son journal).
 */

// ───────────────────────────── Configuration ─────────────────────────────
var CFG = {
  TZ: 'America/Toronto',
  CAL_ID: '1Ur9Kd_wEPs-ERF8cXhqZFm292eFSiot5wp7wrU7zQGk',          // Calendrier de culture
  CLIMAT_FOLDER: '1fmk8svBfc1J_DjwGVBXxbn6gsqxJa9J6',               // 💧 Climat Inkbird — Capteurs & rapports
  RAPPORTS_FOLDER: '1VDR6A3ndEP8M353nmagXyV8qveWiZfbC',             // Rapports climat
  EXPORT_SUBFOLDER: 'Registres des réglages climat',
  SALLES: ['Salle 2', 'Salle 3', 'Salle 4', 'Salle 5', 'Salle 6', 'Salle 7'],
  SALLE_MERES: 'Salle 2',
  // Repli si l'onglet « Registres » du calendrier est illisible
  REGISTRES: {
    'Salle 2': '1_KovE1URLcMQJ7xtHcZFzO9uFSJxFHsCZVi_ngiE-mw',
    'Salle 3': '1gJSXFbsYUUlaxYru9Y-NZ6SdiFvl8vQxOpkBYAvaMZs',
    'Salle 4': '10gpz6q22djraQ1c6gHh2KZ4OGBhjrvU66-tsvvviXaA',
    'Salle 5': '1E9hZBPTxr33lgVSiWQdVXEI_5CcIpn9wGvgAz2z9x5k',
    'Salle 6': '1eF-ob8GoWeETfXG-ABhzNrqAhyEy1lzIlx9zxhXLbYA',
    'Salle 7': '1sxptVloSWaW33zs6mTyKHfT806EW8DiDrcaZgVncCeg'
  },
  // Tolérances (identiques à la tâche « Analyse climat Inkbird »)
  TOL: { T: 1, HR: 5, DIF: 1.5, MARGE: 3 },
  // Tolérances de saisie : écart entre la consigne réglée et la consigne attendue
  TOL_REGLAGE: { T: 0.3, HR: 1.5, CO2: 60, RAMPE: 0.25 },
  CACHE_S: 600
};

var SEMAINES = ['Veg 1', 'Veg 2', 'Flo 1', 'Flo 2', 'Flo 3', 'Flo 4', 'Flo 5', 'Flo 6', 'Flo 7', 'Flo 8', 'Flo 9', 'Flo 10'];

var TABS = {
  'Réglages': ['ID', 'Horodatage', 'Date effective', 'Heure', 'Salle', 'Round', 'Phase', 'Semaine',
    'T jour réglée (°C)', 'HR jour réglée (%)', 'T nuit réglée (°C)', 'HR nuit réglée (%)', 'CO₂ réglé (ppm)', 'Rampe réglée (h)',
    'Cible T jour', 'Cible HR jour', 'Cible T nuit', 'Cible HR nuit', 'Cible CO₂', 'Cible rampe',
    'Attendu thermostat T jour', 'Attendu thermostat T nuit', 'Source de la cible',
    'Conformité du réglage', 'Écarts', 'Note', 'Saisi par', 'Statut', 'Motif annulation', 'Annulé le', 'Tâche calendrier'],
  'Relevés IR': ['ID', 'Horodatage', 'Date', 'Heure', 'Salle', 'Round', 'Semaine', 'Minutes après OFF',
    'T air (°C)', 'HR (%)', 'IR 1 (°C)', 'IR 2 (°C)', 'IR 3 (°C)', 'T feuille moy. (°C)', 'Delta air − feuille (°C)',
    'Pt rosée (°C)', 'Marge feuille − pt rosée (°C)', 'Verdict', 'Note', 'Saisi par', 'Statut', 'Motif annulation', 'Annulé le'],
  'Transition OFF': ['ID', 'Horodatage', 'Date', 'Salle', 'Round', 'Semaine',
    'T pré-OFF', 'HR pré-OFF', 'T +30', 'HR +30', 'T +90', 'HR +90', 'HR max 0–90', 'Minutes > plafond', 'Surface min (°C)',
    'DP pré-OFF', 'DP +30', 'DP +90', 'Chute DP +30', 'Chute DP +90', 'Plafond HR', 'Chute DP cible',
    'Départ DP', 'DP à 90 min', 'Exposition HR', 'Condensation', 'Note', 'Saisi par', 'Statut', 'Motif annulation', 'Annulé le'],
  'Observations': ['ID', 'Horodatage', 'Date', 'Salle', 'Round', 'Semaine', 'Catégorie', 'Observation',
    'Saisi par', 'Statut', 'Motif annulation', 'Annulé le'],
  'Rappels': ['ID', 'Horodatage', 'Salle', 'Round', 'Semaine', 'Date du changement', 'Action', 'Note'],
  'Plan': ['Clé', 'Valeur', 'Mis à jour'],
  'Journal': ['Horodatage', 'Action', 'Objet', 'ID', 'Détails']
};

// ───────────────────────────── Installation ─────────────────────────────
/** À exécuter une fois depuis l'éditeur : crée la feuille et le code d'accès. */
function setup() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  var ss;
  if (id) { ss = SpreadsheetApp.openById(id); }
  else {
    ss = SpreadsheetApp.create('Climat — Amaltis (registre des réglages)');
    try { DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(CFG.CLIMAT_FOLDER)); } catch (e) {}
    props.setProperty('SHEET_ID', ss.getId());
  }
  ss.setSpreadsheetTimeZone(CFG.TZ);
  ensureTabs_(ss);
  if (!props.getProperty('CODE')) {
    var code = String(Math.floor(100000 + Math.random() * 900000));
    props.setProperty('CODE', code);
  }
  Logger.log('Feuille : ' + ss.getUrl());
  Logger.log('Code d\'accès : ' + props.getProperty('CODE'));
  majPlan_(true);
}

/** Pour changer le code : modifier la valeur ci-dessous, exécuter, puis remettre ''. */
function changerCode() {
  var nouveau = '';
  if (!nouveau) throw new Error('Inscrivez le nouveau code dans la fonction changerCode().');
  PropertiesService.getScriptProperties().setProperty('CODE', nouveau);
}

function ensureTabs_(ss) {
  Object.keys(TABS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    var h = TABS[name];
    var cur = sh.getRange(1, 1, 1, h.length).getValues()[0];
    if (cur.join('|') !== h.join('|')) {
      sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#e6ebf0');
      sh.setFrozenRows(1);
    }
  });
  var s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Feuille 1') || ss.getSheetByName('Feuille1');
  if (s1 && ss.getSheets().length > 1) ss.deleteSheet(s1);
}

function ss_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Installation incomplète : exécutez setup() dans l\'éditeur Apps Script.');
  return SpreadsheetApp.openById(id);
}

// ───────────────────────────── API ─────────────────────────────
function doGet() {
  return json_({ ok: true, app: 'Climat — Amaltis', info: 'API active. Utilisez l\'application.' });
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, erreur: 'Requête illisible.' }); }
  try {
    var props = PropertiesService.getScriptProperties();
    if (!verifierCode_(req.code, props)) return json_({ ok: false, auth: false, erreur: 'Code invalide.' });
    var a = req.action, p = req.p || {};
    var actions = {
      ping: function () { return { ok: true }; },
      accueil: function () { return accueil_(p.force); },
      salle: function () { return salle_(p.salle, p.round, p.force); },
      reglage: function () { return enregistrerReglage_(p); },
      ir: function () { return enregistrerIR_(p); },
      transition: function () { return enregistrerTransition_(p); },
      observation: function () { return enregistrerObservation_(p); },
      annuler: function () { return annuler_(p); },
      rappel: function () { return actionRappel_(p); },
      exporter: function () { return exporter_(p); },
      plan: function () { return { plan: majPlan_(!!p.force) }; },
      sommaireIR: function () { return sommaireIR_(); },
      cible: function () { return modifierCible_(p); },
      descenteApercu: function () { return descente_(p, false); },
      descente: function () { return descente_(p, true); },
      historique: function () { return historique_(p.salle); }
    };
    if (!actions[a]) return json_({ ok: false, erreur: 'Action inconnue : ' + a });
    var out = actions[a]();
    out.ok = true;
    return json_(out);
  } catch (err) {
    journal_('erreur', req && req.action, '', String(err && err.stack || err).slice(0, 500));
    return json_({ ok: false, erreur: String(err && err.message || err) });
  }
}

function verifierCode_(code, props) {
  var cache = CacheService.getScriptCache();
  var echecs = Number(cache.get('echecs_v2') || 0);
  if (echecs >= 10) return false; // 10 essais ratés → pause de 10 minutes
  var ok = code && String(code).trim() === String(props.getProperty('CODE'));
  if (!ok) cache.put('echecs_v2', String(echecs + 1), 600);
  else if (echecs) cache.remove('echecs_v2');
  return !!ok;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────────── Dates ─────────────────────────────
function aujourdhui_() { return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd'); }
function maintenant_() { return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd HH:mm'); }
function heure_() { return Utilities.formatDate(new Date(), CFG.TZ, 'HH:mm'); }
function dIso_(iso) { var p = iso.split('-'); return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2], 12)); }
function isoD_(d) { return d.toISOString().slice(0, 10); }
function plusJours_(iso, n) { var d = dIso_(iso); d.setUTCDate(d.getUTCDate() + n); return isoD_(d); }
function ecartJours_(a, b) { return Math.round((dIso_(b) - dIso_(a)) / 86400000); }
function jourSemaine_(iso) { return dIso_(iso).getUTCDay(); } // 0 = dimanche

/** Date d'une cellule de registre (objet Date ou « JJ/MM/AAAA ») → AAAA-MM-JJ. */
function dateCellule_(v, tz) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz || CFG.TZ, 'yyyy-MM-dd');
  var s = String(v || '').trim();
  var m = s.match(/^(\d{1,2})[\/\.-](\d{1,2})[\/\.-](\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[0];
  return null;
}

/** Jour ouvrable conseillé : samedi → vendredi, dimanche → lundi (règle du calendrier). */
function dateConseillee_(iso) {
  var j = jourSemaine_(iso);
  if (j === 6) return plusJours_(iso, -1);
  if (j === 0) return plusJours_(iso, 1);
  return iso;
}

function num_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  var s = String(v).replace(/\s/g, '').replace(',', '.');
  var n = parseFloat(s);
  return isFinite(n) ? n : null;
}
function arr1_(n) { return n === null || n === undefined ? null : Math.round(n * 10) / 10; }
function arr2_(n) { return n === null || n === undefined ? null : Math.round(n * 100) / 100; }

// ───────────────────────────── Psychrométrie ─────────────────────────────
function svp_(t) { return 0.6108 * Math.exp(17.27 * t / (t + 237.3)); }
function ptRosee_(t, hr) {
  if (t === null || hr === null || hr <= 0) return null;
  var a = Math.log(hr / 100) + 17.27 * t / (t + 237.3);
  return 237.3 * a / (17.27 - a);
}
function vpdFeuille_(tAir, hr, tFeuille) {
  if (tAir === null || hr === null) return null;
  return svp_(tFeuille) - svp_(tAir) * hr / 100;
}

// ───────────────────────────── Arrimage des semaines ─────────────────────────────
/**
 * Semaine du plan pour un jour donné d'un round (identique à l'analyse Inkbird) :
 *  VEG : Veg 1, sauf les 7 derniers jours avant FLO (Veg 2).
 *  FLO : Flo 1–3 depuis FLO J1 ; Flo 7–10 à rebours depuis le dernier jour de FLO ;
 *        le milieu prend Flo 4, 5, 6 dans l'ordre ; une semaine partielle est
 *        rattachée à la semaine voisine (la dernière semaine du milieu).
 */
function semaineDuJour_(veg, flo, jour) {
  if (jour < 1) return null;
  if (jour <= veg) return { label: jour > veg - 7 ? 'Veg 2' : 'Veg 1', phase: 'VEG', jourPhase: jour };
  var f = jour - veg;
  if (f > flo) return null;
  var res = { phase: 'FLO', jourPhase: f };
  if (f <= 21) { res.label = 'Flo ' + Math.ceil(f / 7); return res; }
  if (f > flo - 28) { res.label = 'Flo ' + (10 - Math.floor((flo - f) / 7)); return res; }
  // Milieu comprimé : les jours entre Flo 3 et la descente sont répartis entre Flo 4, 5 et 6
  // (70 j → 7/7/7 ; 65 j → 5/5/6 ; 63 j → 4/5/5 ; 56 j → 2/2/3). Aucune semaine n'est sautée.
  var m = flo - 49, k = f - 22;
  var s4 = Math.floor(m / 3), s5 = Math.floor((m + 1) / 3);
  res.label = k < s4 ? 'Flo 4' : k < s4 + s5 ? 'Flo 5' : 'Flo 6';
  return res;
}

// ───────────────────────────── Paramètres du calendrier ─────────────────────────────
function paramsCalendrier_() {
  var c = CacheService.getScriptCache(), k = 'calp';
  var hit = c.get(k); if (hit) return JSON.parse(hit);
  var out = { offJour: -1, offNuit: 0, registres: {} };
  try {
    var cal = SpreadsheetApp.openById(CFG.CAL_ID);
    var v = cal.getSheetByName('Paramètres').getDataRange().getValues();
    v.forEach(function (r) {
      var nom = String(r[0] || '');
      if (/décalage thermostat jour/i.test(nom) && num_(r[1]) !== null) out.offJour = num_(r[1]);
      if (/décalage thermostat nuit/i.test(nom) && num_(r[1]) !== null) out.offNuit = num_(r[1]);
    });
    var reg = cal.getSheetByName('Registres').getDataRange().getValues();
    reg.slice(1).forEach(function (r) { if (r[0] && r[1] && String(r[2]).toLowerCase() !== 'non') out.registres[String(r[0]).trim()] = String(r[1]).trim(); });
  } catch (e) { out.erreur = String(e.message || e); }
  CFG.SALLES.forEach(function (s) { if (!out.registres[s]) out.registres[s] = CFG.REGISTRES[s]; });
  c.put(k, JSON.stringify(out), 3600);
  return out;
}

// ───────────────────────────── Plan VPD (dernière version) ─────────────────────────────
function majPlan_(force) {
  var c = CacheService.getScriptCache();
  if (!force) { var hit = c.get('plan'); if (hit) return JSON.parse(hit); }
  var sh = ss_().getSheetByName('Plan');
  var stocke = null;
  var v = sh.getRange('B2').getValue();
  if (v) { try { stocke = JSON.parse(v); } catch (e) {} }

  // Recherche de la version la plus récente « Plan_VPD_vNN »
  // Candidats : « Plan_VPD_vNN.xlsx » ou Google Sheet du même nom ; on écarte les
  // fichiers de verrouillage d'Excel (« ~$Plan_VPD_v55.xlsx ») et les fichiers minuscules.
  var cands = [];
  var it = DriveApp.searchFiles("title contains 'Plan_VPD_v' and trashed = false");
  while (it.hasNext()) {
    var f = it.next();
    var m = f.getName().match(/^Plan_VPD_v(\d+)(\.xlsx)?$/i);
    if (!m) continue;
    var gs = f.getMimeType() === MimeType.GOOGLE_SHEETS;
    if (!gs && (!/\.xlsx$/i.test(f.getName()) || f.getSize() < 5000)) continue;
    cands.push({ n: +m[1], f: f, t: f.getLastUpdated().getTime() });
  }
  cands.sort(function (a, b) { return b.n - a.n || b.t - a.t; });
  if (!cands.length) { if (stocke) { c.put('plan', JSON.stringify(stocke), 3600); return stocke; } throw new Error('Aucun fichier Plan_VPD_vNN trouvé dans Drive.'); }
  var meilleur = cands[0];
  var modif = meilleur.f.getLastUpdated().toISOString();
  if (stocke && stocke.fileId === meilleur.f.getId() && stocke.modifie === modif && !force) {
    c.put('plan', JSON.stringify(stocke), 3600); return stocke;
  }
  var plan = null, erreurs = [];
  for (var ci = 0; ci < cands.length && !plan; ci++) {
    try { plan = lirePlanFichier_(cands[ci].f); meilleur = cands[ci]; modif = meilleur.f.getLastUpdated().toISOString(); }
    catch (e) { erreurs.push(cands[ci].f.getName() + ' : ' + e.message); }
    if (!plan && stocke && stocke.fileId === cands[ci].f.getId()) break;
  }
  if (!plan) {
    if (stocke) { stocke.avertissement = 'Lecture du plan impossible (' + erreurs.join(' ; ') + ') ; version en cache utilisée.'; return stocke; }
    throw new Error('Lecture du plan impossible : ' + erreurs.join(' ; '));
  }
  if (meilleur !== cands[0]) plan.avertissement = 'Version la plus récente illisible (' + erreurs.join(' ; ') + ') ; ' + meilleur.f.getName() + ' utilisée.';
  plan.version = 'v' + meilleur.n;
  plan.nom = meilleur.f.getName();
  plan.fileId = meilleur.f.getId();
  plan.modifie = modif;
  plan.lu = maintenant_();
  sh.getRange('A2:C2').setValues([['plan', JSON.stringify(plan), plan.lu]]);
  c.put('plan', JSON.stringify(plan), 3600);
  journal_('plan', 'Plan VPD', plan.fileId, 'Lecture de ' + plan.nom);
  return plan;
}

function lirePlanFichier_(file) {
  var tmpId = null, book;
  if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
    book = SpreadsheetApp.openById(file.getId());
  } else {
    // Conversion xlsx → Google Sheets temporaire (API Drive v3), lecture, puis corbeille.
    var r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + file.getId() + '/copy?fields=id', {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ name: 'tmp_climat_' + file.getName(), mimeType: 'application/vnd.google-apps.spreadsheet' }),
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
    });
    if (r.getResponseCode() !== 200) throw new Error('conversion refusée (' + r.getResponseCode() + ') ' + r.getContentText().slice(0, 160));
    tmpId = JSON.parse(r.getContentText()).id;
    book = SpreadsheetApp.openById(tmpId);
  }
  try {
    var plan = { semaines: {}, transition: {}, params: { deltaJour: 1.5, margeRosee: 3, baisseDp30: 0.3, toleranceDp90: 0.2 } };
    // Onglet « Plan VPD »
    var pv = book.getSheetByName('Plan VPD').getDataRange().getValues();
    var hr = -1;
    for (var i = 0; i < pv.length; i++) if (String(pv[i][0]).trim() === 'Semaine') { hr = i; break; }
    if (hr < 0) throw new Error('en-tête « Semaine » introuvable dans Plan VPD');
    var H = pv[hr].map(function (h) { return String(h).replace(/\s+/g, ' ').trim(); });
    var col = function (re) { for (var j = 0; j < H.length; j++) if (re.test(H[j])) return j; return -1; };
    var C = {
      phase: col(/^Phase$/), tj: col(/^T air jour/), hrj: col(/^HR jour/), tn: col(/^T air nuit/), hrn: col(/^HR nuit/),
      dif: col(/^DIF/), rampe: col(/^Rampe/), co2: col(/^CO₂ jour/), vmin: col(/^Cible min j/), vmax: col(/^Cible max j/), hrnmax: col(/^HR max n/)
    };
    for (var k = hr + 1; k < pv.length; k++) {
      var s = String(pv[k][0]).trim();
      if (SEMAINES.indexOf(s) < 0) continue;
      var g = function (j) { return j < 0 ? null : num_(pv[k][j]); };
      plan.semaines[s] = {
        phase: C.phase >= 0 ? String(pv[k][C.phase]) : '', tj: g(C.tj), hrj: g(C.hrj), tn: g(C.tn), hrn: g(C.hrn),
        dif: g(C.dif), rampe: g(C.rampe), co2: g(C.co2), vpdMin: g(C.vmin), vpdMax: g(C.vmax), hrnMax: g(C.hrnmax)
      };
    }
    // Onglet « Transition OFF » : plafonds HR et chute DP cible
    var to = book.getSheetByName('Transition OFF');
    if (to) {
      var tv = to.getDataRange().getValues(), th = -1;
      for (var a = 0; a < tv.length; a++) if (String(tv[a][0]).trim() === 'Semaine') { th = a; break; }
      if (th >= 0) {
        var TH = tv[th].map(function (h) { return String(h).replace(/\s+/g, ' ').trim(); });
        var cp = TH.findIndex(function (h) { return /^Plafond RH/.test(h); });
        var cc = TH.findIndex(function (h) { return /^Chute DP cible/.test(h); });
        for (var b = th + 1; b < tv.length; b++) {
          var sw = String(tv[b][0]).trim();
          if (SEMAINES.indexOf(sw) < 0) { if (sw === 'Date') break; continue; }
          plan.transition[sw] = { plafond: num_(tv[b][cp]), chuteCible: num_(tv[b][cc]) };
        }
      }
    }
    // Onglet « PARAMS »
    var pa = book.getSheetByName('PARAMS');
    if (pa) pa.getDataRange().getValues().forEach(function (r) {
      var n = String(r[0]);
      if (/Delta feuille \/ air — JOUR/i.test(n) && num_(r[1]) !== null) plan.params.deltaJour = num_(r[1]);
      if (/Marge minimale point de rosée/i.test(n) && num_(r[1]) !== null) plan.params.margeRosee = num_(r[1]);
      if (/Baisse DP minimale après 30/i.test(n) && num_(r[1]) !== null) plan.params.baisseDp30 = num_(r[1]);
      if (/Tolérance cible DP à 90/i.test(String(r[3])) && num_(r[4]) !== null) plan.params.toleranceDp90 = num_(r[4]);
      if (/Preset de bandes/i.test(n)) plan.params.preset = String(r[1]);
    });
    if (Object.keys(plan.semaines).length < 12) throw new Error('semaines incomplètes dans Plan VPD');
    return plan;
  } finally {
    if (tmpId) try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (e) {}
  }
}

// ───────────────────────────── Registres de culture ─────────────────────────────
/** Lit les rounds d'un registre (les deux plus récents, avec leurs jours). */
function lireRegistre_(salle, force) {
  var c = CacheService.getScriptCache(), k = 'reg:' + salle;
  if (!force) { var hit = c.get(k); if (hit) return JSON.parse(hit); }
  var id = paramsCalendrier_().registres[salle];
  var out = { salle: salle, fichierId: id, rounds: [], erreur: null };
  if (salle === CFG.SALLE_MERES) { out.continu = true; c.put(k, JSON.stringify(out), CFG.CACHE_S); return out; }
  try {
    var book = SpreadsheetApp.openById(id);
    var tz = book.getSpreadsheetTimeZone();
    var tabs = book.getSheets().filter(function (sh) { return /round/i.test(sh.getName()); }).map(function (sh) {
      var b = sh.getRange('B1:B3').getDisplayValues(); // texte affiché : évite les décalages de fuseau (registre en Etc/GMT)
      return { sh: sh, start: dateCellule_(b[0][0], tz), veg: num_(b[1][0]), flo: num_(b[2][0]) };
    }).filter(function (t) { return t.start && t.veg !== null && t.flo !== null; });
    tabs.sort(function (a, b) { return a.start < b.start ? -1 : 1; });
    tabs.slice(-2).forEach(function (t) { out.rounds.push(lireRound_(t, tz)); });
  } catch (e) { out.erreur = String(e.message || e); }
  c.put(k, JSON.stringify(out), CFG.CACHE_S);
  return out;
}

function lireRound_(t, tz) {
  var sh = t.sh, nom = sh.getName().trim();
  var r = { onglet: nom, round: (nom.match(/(\d+)/) || [, nom])[1], start: t.start, veg: t.veg, flo: t.flo,
    fin: plusJours_(t.start, t.veg + t.flo - 1), recolte: plusJours_(t.start, t.veg + t.flo), table: null, mesures: {}, gacp: false };
  var lastRow = Math.max(7, sh.getLastRow()), lastCol = Math.max(10, sh.getLastColumn());
  var v = sh.getRange(1, 1, Math.min(lastRow, 200), lastCol).getValues();
  // Tableau de consignes par semaine (format GACP) : en-tête « Semaine » ligne 5, colonnes ≥ U
  for (var j = 20; j < lastCol; j++) {
    if (String(v[4][j]).trim() === 'Semaine' && String(v[4][j + 1]).trim() === 'Phase') {
      r.table = {}; r.gacp = true;
      for (var i = 5; i < Math.min(v.length, 30); i++) {
        var s = String(v[i][j]).trim();
        if (SEMAINES.indexOf(s) < 0) continue;
        var g = function (o) { return num_(v[i][j + o]); };
        r.table[s] = { phase: String(v[i][j + 1]), tj: g(2), hrj: g(3), tn: g(4), hrn: g(5), vpdMin: g(6), vpdMax: g(7), hrnMax: g(8), dif: g(9), rampe: g(10), co2: g(11) };
      }
      break;
    }
  }
  // Mesures Inkbird reportées au registre (moyennes jour/nuit), par en-tête de la ligne 6
  var H = v[5].map(function (h) { return String(h).replace(/\s+/g, ' ').trim(); });
  var ci = function (re) { for (var q = 0; q < H.length; q++) if (re.test(H[q])) return q; return -1; };
  var cDay = H.indexOf('Day'), cTj = ci(/^T air jour.*Inkbird/i), cHj = ci(/^HR jour.*Inkbird/i),
    cTn = ci(/^T air nuit.*Inkbird/i), cHn = ci(/^HR nuit.*Inkbird/i), cCo = ci(/^CO₂ \(ppm\) mesuré/i);
  if (cDay >= 0 && (cTj >= 0 || cTn >= 0)) {
    for (var row = 6; row < v.length; row++) {
      var d = num_(v[row][cDay]);
      if (!d) continue;
      var m = { tj: cTj >= 0 ? num_(v[row][cTj]) : null, hrj: cHj >= 0 ? num_(v[row][cHj]) : null,
        tn: cTn >= 0 ? num_(v[row][cTn]) : null, hrn: cHn >= 0 ? num_(v[row][cHn]) : null, co2: cCo >= 0 ? num_(v[row][cCo]) : null };
      if (m.tj !== null || m.tn !== null || m.hrj !== null || m.hrn !== null) r.mesures[plusJours_(t.start, d - 1)] = m;
    }
  }
  return r;
}

// ───────────────────────────── Cibles ─────────────────────────────
function cible_(round, semaine, plan) {
  var p = plan.semaines[semaine] || null;
  var t = round && round.table && round.table[semaine];
  var src = t ? 'Registre (' + round.onglet + ')' : 'Plan VPD ' + plan.version;
  var base = t || p;
  if (!base) return null;
  var c = { semaine: semaine, phase: base.phase, tj: base.tj, hrj: base.hrj, tn: base.tn, hrn: base.hrn, co2: base.co2, rampe: base.rampe,
    vpdMin: base.vpdMin, vpdMax: base.vpdMax, hrnMax: base.hrnMax, source: src, divergences: [] };
  c.dif = c.tj !== null && c.tn !== null ? arr1_(c.tj - c.tn) : null;
  var delta = plan.params.deltaJour;
  c.vpd = c.tj !== null && c.hrj !== null ? arr2_(vpdFeuille_(c.tj, c.hrj, c.tj - delta)) : null;
  c.ptRoseeNuit = arr1_(ptRosee_(c.tn, c.hrn));
  var tr = plan.transition[semaine] || {};
  c.plafondOff = tr.plafond; c.chuteDpCible = arr1_(tr.chuteCible);
  if (t && p) {
    [['tj', 'T jour'], ['hrj', 'HR jour'], ['tn', 'T nuit'], ['hrn', 'HR nuit'], ['co2', 'CO₂'], ['rampe', 'Rampe']].forEach(function (f) {
      if (t[f[0]] !== null && p[f[0]] !== null && Math.abs(t[f[0]] - p[f[0]]) > 0.01) c.divergences.push(f[1] + ' : registre ' + t[f[0]] + ' / plan ' + p[f[0]]);
    });
  }
  return c;
}

function attendu_(c, cal) {
  if (!c) return null;
  return { tj: c.tj === null ? null : arr1_(c.tj + cal.offJour), tn: c.tn === null ? null : arr1_(c.tn + cal.offNuit), hrj: c.hrj, hrn: c.hrn, co2: c.co2, rampe: c.rampe };
}

/** Situation d'une salle à une date : round, semaine, cible. */
function situation_(reg, iso, plan) {
  if (reg.continu) {
    return { round: null, onglet: 'Continu', phase: 'VEG', semaine: 'Veg 1', jour: null, jourPhase: null, vide: false, continu: true };
  }
  var r = null;
  for (var i = reg.rounds.length - 1; i >= 0; i--) if (reg.rounds[i].start <= iso) { r = reg.rounds[i]; break; }
  if (!r) return { vide: true, motif: 'Aucun round en cours' };
  var jour = ecartJours_(r.start, iso) + 1;
  var s = semaineDuJour_(r.veg, r.flo, jour);
  if (!s) return { vide: true, motif: 'Salle vide depuis la récolte du ' + r.recolte, roundObj: r, onglet: r.onglet, round: r.round };
  return { roundObj: r, round: r.round, onglet: r.onglet, phase: s.phase, semaine: s.label, jour: jour, jourPhase: s.jourPhase, vide: false };
}

/** Changements de consigne d'un round (jours où la semaine du plan change). */
function changements_(round, plan) {
  var out = [], prev = null, prevC = null;
  var total = round.veg + round.flo;
  for (var j = 1; j <= total; j++) {
    var s = semaineDuJour_(round.veg, round.flo, j).label;
    if (s === prev) continue;
    var iso = plusJours_(round.start, j - 1);
    var c = cible_(round, s, plan);
    var diff = [];
    if (prevC && c) [['tj', 'T jour'], ['hrj', 'HR jour'], ['tn', 'T nuit'], ['hrn', 'HR nuit'], ['co2', 'CO₂'], ['rampe', 'Rampe']].forEach(function (f) {
      if (c[f[0]] !== prevC[f[0]]) diff.push({ champ: f[1], de: prevC[f[0]], a: c[f[0]] });
    });
    out.push({ semaine: s, de: prev, date: iso, conseillee: dateConseillee_(iso), jour: j, debut: !prev, modifs: diff, sansChangement: !!prev && diff.length === 0 });
    prev = s; prevC = c;
  }
  return out;
}

// ───────────────────────────── Feuilles de l'app ─────────────────────────────
function lireOnglet_(nom) {
  var sh = ss_().getSheetByName(nom);
  var v = sh.getDataRange().getValues();
  var h = v[0];
  return v.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    var o = {};
    h.forEach(function (k, i) { var x = r[i]; o[k] = x instanceof Date ? Utilities.formatDate(x, CFG.TZ, 'yyyy-MM-dd HH:mm') : x; });
    return o;
  });
}

function reglagesActifs_(salle) {
  return lireOnglet_('Réglages').filter(function (r) { return r['Statut'] !== 'annulé' && (!salle || r['Salle'] === salle); })
    .map(function (r) { r['Date effective'] = String(r['Date effective']).slice(0, 10); return r; })
    .sort(function (a, b) { var x = a['Date effective'] + ' ' + a['Heure'], y = b['Date effective'] + ' ' + b['Heure']; return x < y ? -1 : x > y ? 1 : (a['Horodatage'] < b['Horodatage'] ? -1 : 1); });
}

/** Réglage en vigueur à une date (dernier réglage actif à cette date ou avant). */
function reglageA_(liste, iso) {
  var r = null;
  for (var i = 0; i < liste.length; i++) if (liste[i]['Date effective'] <= iso) r = liste[i];
  return r;
}

function regVal_(r) {
  if (!r) return null;
  return { tj: num_(r['T jour réglée (°C)']), hrj: num_(r['HR jour réglée (%)']), tn: num_(r['T nuit réglée (°C)']), hrn: num_(r['HR nuit réglée (%)']),
    co2: num_(r['CO₂ réglé (ppm)']), rampe: num_(r['Rampe réglée (h)']) };
}

/** Conformité d'un réglage par rapport à la consigne attendue au thermostat. */
function verifierReglage_(val, att) {
  if (!val || !att) return { ok: null, ecarts: [] };
  var T = CFG.TOL_REGLAGE, e = [];
  var chk = function (k, nom, tol, unite) {
    if (val[k] === null || val[k] === undefined || att[k] === null || att[k] === undefined) return;
    var d = val[k] - att[k];
    if (Math.abs(d) > tol) e.push(nom + ' ' + val[k] + unite + ' au lieu de ' + att[k] + unite + ' (' + (d > 0 ? '+' : '') + arr1_(d) + ')');
  };
  chk('tj', 'T jour', T.T, ' °C'); chk('tn', 'T nuit', T.T, ' °C'); chk('hrj', 'HR jour', T.HR, ' %'); chk('hrn', 'HR nuit', T.HR, ' %');
  chk('co2', 'CO₂', T.CO2, ' ppm'); chk('rampe', 'Rampe', T.RAMPE, ' h');
  if (val.tj !== null && val.tn !== null && val.tn > val.tj) e.unshift('Nuit plus chaude que le jour (DIF inversé)');
  return { ok: e.length === 0, ecarts: e };
}

/** Diagnostic d'un jour : cible / réglé / mesuré. */
function diagnostic_(c, att, reg, mes) {
  var d = { code: '', texte: '', details: [] };
  if (!reg) { d.code = 'non-inscrit'; d.texte = 'Réglage non inscrit'; }
  else {
    var v = verifierReglage_(regVal_(reg), att);
    if (v.ok === false) { d.code = 'mauvais-reglage'; d.texte = 'Mauvais réglage'; d.details = v.ecarts; }
  }
  if (mes && c) {
    var T = CFG.TOL, e = [];
    if (mes.tj !== null && c.tj !== null && Math.abs(mes.tj - c.tj) > T.T) e.push('T jour ' + mes.tj + ' vs ' + c.tj);
    if (mes.tn !== null && c.tn !== null && Math.abs(mes.tn - c.tn) > T.T) e.push('T nuit ' + mes.tn + ' vs ' + c.tn);
    if (mes.hrj !== null && c.hrj !== null && Math.abs(mes.hrj - c.hrj) > T.HR) e.push('HR jour ' + mes.hrj + ' vs ' + c.hrj);
    if (mes.hrn !== null && c.hrnMax !== null && mes.hrn > c.hrnMax) e.push('HR nuit ' + mes.hrn + ' > max ' + c.hrnMax);
    else if (mes.hrn !== null && c.hrn !== null && Math.abs(mes.hrn - c.hrn) > T.HR) e.push('HR nuit ' + mes.hrn + ' vs ' + c.hrn);
    if (mes.tj !== null && mes.tn !== null && c.dif !== null && Math.abs((mes.tj - mes.tn) - c.dif) > T.DIF) e.push('DIF ' + arr1_(mes.tj - mes.tn) + ' vs ' + c.dif);
    if (!d.code) {
      if (e.length) { d.code = 'ne-suit-pas'; d.texte = 'La salle ne suit pas le réglage'; d.details = e; }
      else { d.code = 'conforme'; d.texte = 'Conforme'; }
    } else if (e.length) d.mesure = e;
  } else if (!d.code) { d.code = 'sans-mesure'; d.texte = 'Réglage conforme · pas de mesure'; }
  return d;
}

// ───────────────────────────── Écran d'accueil ─────────────────────────────
function accueil_(force) {
  var plan = majPlan_(false), cal = paramsCalendrier_(), today = aujourdhui_();
  if (force) { CacheService.getScriptCache().removeAll(CFG.SALLES.map(function (s) { return 'reg:' + s; }).concat(['calp', 'tachesClimat'], CFG.SALLES.map(function (s) { return 'rap:' + s; }))); cal = paramsCalendrier_(); }
  var reglages = reglagesActifs_(null), rappels = lireOnglet_('Rappels'), taches = tachesClimat_();
  var salles = CFG.SALLES.map(function (s) { return resumeSalle_(s, lireRegistre_(s, force), plan, cal, today, reglages, rappels, taches); });
  return { aujourdhui: today, plan: { version: plan.version, nom: plan.nom, lu: plan.lu, avertissement: plan.avertissement || null },
    offsets: { jour: cal.offJour, nuit: cal.offNuit }, salles: salles };
}

function resumeSalle_(salle, reg, plan, cal, today, reglagesTous, rappels, taches) {
  var o = { salle: salle, erreur: reg.erreur };
  var sit = situation_(reg, today, plan);
  o.situation = { round: sit.round, onglet: sit.onglet, phase: sit.phase, semaine: sit.semaine, jour: sit.jour, jourPhase: sit.jourPhase, vide: sit.vide, motif: sit.motif || null, continu: !!sit.continu };
  var r = sit.roundObj || null;
  if (r) o.round = { onglet: r.onglet, start: r.start, veg: r.veg, flo: r.flo, recolte: r.recolte, floJ1: plusJours_(r.start, r.veg), gacp: r.gacp };
  var regs = reglagesTous.filter(function (x) { return x['Salle'] === salle; });
  var dernier = regs.length ? regs[regs.length - 1] : null;
  o.dernierReglage = dernier ? resumeReglage_(dernier) : null;
  if (!sit.vide) {
    var c = cible_(r, sit.semaine, plan);
    o.cible = c; o.attendu = attendu_(c, cal);
    var verif = verifierReglage_(regVal_(dernier), o.attendu);
    o.etatReglage = dernier ? (verif.ok ? 'conforme' : 'ecart') : 'aucun';
    o.ecartsReglage = verif.ecarts;
  }
  // Rappels
  if (r && !sit.continu) {
    var ch = changements_(r, plan);
    var conf = function (e) { return regs.some(function (x) { return x['Round'] == r.round && SEMAINES.indexOf(x['Semaine']) >= SEMAINES.indexOf(e.semaine) && x['Date effective'] >= plusJours_(e.conseillee, -3); }) ||
      rappels.some(function (x) { return x['Salle'] === salle && x['Round'] == r.round && x['Semaine'] === e.semaine; }); };
    var passes = ch.filter(function (e) { return e.conseillee <= today && !e.sansChangement; });
    var enAttente = passes.length ? passes[passes.length - 1] : null;
    if (enAttente && !conf(enAttente)) {
      enAttente.retard = ecartJours_(enAttente.conseillee, today);
      o.rappel = enAttente;
    }
    o.aujourdhuiChangement = ch.some(function (e) { return e.conseillee === today && !e.sansChangement; });
    o.prochain = ch.filter(function (e) { return e.conseillee > today && !e.sansChangement; })[0] || null;
    // Plan de rattrapage (récolte avancée) : il prime sur la consigne de la semaine tant qu'il dure
    var pa = planActif_(salle, r.round);
    if (pa) {
      o.plan = pa;
      var step = pa.etapes.filter(function (e) { return e.date <= today; }).pop() || null;
      o.planProchaine = pa.etapes.filter(function (e) { return e.date > today; })[0] || null;
      if (step) {
        o.propose = step.val;
        var vd = regVal_(dernier), vr = verifierReglage_(vd, step.val);
        o.etatReglage = dernier ? (vr.ok ? 'conforme' : 'ecart') : 'aucun';
        o.ecartsReglage = vr.ecarts;
        if (!dernier || !vr.ok) {
          var mods = [];
          ['tj', 'hrj', 'tn', 'hrn', 'co2', 'rampe'].forEach(function (k) { if (!vd || vd[k] !== step.val[k]) mods.push({ champ: CHAMPS[k], de: vd ? vd[k] : null, a: step.val[k] }); });
          o.rappel = { semaine: step.semaine, date: step.date, conseillee: step.date, retard: ecartJours_(step.date, today), rattrapage: true, modifs: mods };
        } else if (o.rappel && o.rappel.conseillee <= step.date) o.rappel = null;
      }
    }
  }
  o.propose = o.propose || o.attendu || null;
  // Tâches Climat du calendrier
  var tc = taches.filter(function (t) { return t.salle === salle; });
  o.tachesCalendrier = tc.filter(function (t) { return t.statut !== 'fait' && t.date <= plusJours_(today, 7); });
  if (r) {
    var chs = changements_(r, plan);
    o.ecartsCalendrier = [];
    tc.forEach(function (t) {
      if (t.statut === 'fait') return;
      var sem = semaineDeTache_(t.titre);
      var e = chs.filter(function (x) { return x.semaine === sem; })[0];
      var prevue = t.prevue || t.date; // date d'origine si la tâche a été reportée
      if (e && prevue >= r.start && prevue <= r.recolte && e.conseillee !== prevue)
        o.ecartsCalendrier.push({ titre: t.titre, calendrier: prevue, plan: e.conseillee, ecart: ecartJours_(e.conseillee, prevue) });
    });
  }
  // Dernière mesure Inkbird au registre
  if (r) {
    // Dernière valeur connue de CHAQUE grandeur : la nuit d'hier compte même si celle d'aujourd'hui manque
    var jours = Object.keys(r.mesures).sort(), m = {}, dates = {};
    ['tj', 'hrj', 'tn', 'hrn', 'co2'].forEach(function (k) {
      for (var q = jours.length - 1; q >= 0; q--) { var v = r.mesures[jours[q]][k]; if (v !== null && v !== undefined) { m[k] = v; dates[k] = jours[q]; break; } }
    });
    if (Object.keys(m).length) o.derniereMesure = { date: dates.tj || dates.tn || jours[jours.length - 1], m: m, dates: dates };
  }
  return o;
}

function resumeReglage_(x) {
  return { id: x['ID'], date: x['Date effective'], heure: x['Heure'], semaine: x['Semaine'], round: x['Round'], val: regVal_(x),
    conformite: x['Conformité du réglage'], ecarts: x['Écarts'], note: x['Note'], horodatage: x['Horodatage'] };
}

// ───────────────────────────── Calendrier ─────────────────────────────
function tachesClimat_() {
  var c = CacheService.getScriptCache(), hit = c.get('tachesClimat');
  if (hit) return JSON.parse(hit);
  var out = [];
  try {
    var sh = SpreadsheetApp.openById(CFG.CAL_ID).getSheetByName('Tâches');
    var v = sh.getDataRange().getValues(), h = v[0];
    var I = function (n) { return h.indexOf(n); };
    var iT = I('Type'), iS = I('Salle'), iD = I('Date'), iTi = I('Titre'), iSt = I('Statut'), iId = I('ID'), iDe = I('Détails'), iP = I('Date prévue');
    for (var i = 1; i < v.length; i++) {
      if (String(v[i][iT]) !== 'Climat') continue;
      out.push({ ligne: i + 1, id: String(v[i][iId]), salle: String(v[i][iS]), date: dateCellule_(v[i][iD]) || String(v[i][iD]).slice(0, 10),
        prevue: iP >= 0 && v[i][iP] ? (dateCellule_(v[i][iP]) || String(v[i][iP]).slice(0, 10)) : null,
        titre: String(v[i][iTi]), statut: String(v[i][iSt]), details: String(v[i][iDe] || '') });
    }
  } catch (e) {}
  c.put('tachesClimat', JSON.stringify(out), 300);
  return out;
}

/** « Climat — Flo S7 » → « Flo 7 » ; « Climat — Veg S2 » → « Veg 2 ». */
function semaineDeTache_(titre) {
  var m = String(titre).match(/(Veg|Flo)\s*S?\s*(\d+)/i);
  return m ? (m[1].charAt(0).toUpperCase() + m[1].slice(1).toLowerCase() + ' ' + m[2]) : null;
}

function cocherTacheCalendrier_(salle, semaine, details) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var cal = SpreadsheetApp.openById(CFG.CAL_ID);
    var sh = cal.getSheetByName('Tâches');
    var v = sh.getDataRange().getValues(), h = v[0];
    var I = function (n) { return h.indexOf(n); };
    var today = aujourdhui_();
    var cands = [];
    for (var i = 1; i < v.length; i++) {
      if (String(v[i][I('Type')]) !== 'Climat' || String(v[i][I('Salle')]) !== salle || String(v[i][I('Statut')]) === 'fait') continue;
      var d = dateCellule_(v[i][I('Date')]) || '';
      if (d > plusJours_(today, 7)) continue;
      var sem = semaineDeTache_(v[i][I('Titre')]);
      cands.push({ i: i, sem: sem, d: d, rang: sem === semaine ? 0 : 1 });
    }
    if (!cands.length) return null;
    cands.sort(function (a, b) { return a.rang - b.rang || (a.d < b.d ? 1 : -1); });
    var t = cands[0], row = t.i + 1, now = maintenant_();
    sh.getRange(row, I('Statut') + 1).setValue('fait');
    sh.getRange(row, I('Fait par') + 1).setValue('Phil');
    sh.getRange(row, I('Fait le') + 1).setValue(now);
    sh.getRange(row, I('Modifié par') + 1).setValue('Phil');
    sh.getRange(row, I('Modifié le') + 1).setValue(now);
    var id = String(v[t.i][I('ID')]), titre = String(v[t.i][I('Titre')]);
    try { cal.getSheetByName('Journal').appendRow([now, 'Phil', 'fait', 'tâche', id, salle + ' · ' + titre + ' · coché depuis l\'app Climat · ' + details]); } catch (e) {}
    CacheService.getScriptCache().remove('tachesClimat');
    return { id: id, titre: titre, date: t.d };
  } finally { lock.releaseLock(); }
}

// ───────────────────────────── Écran d'une salle ─────────────────────────────
function salle_(salle, ongletVoulu, force) {
  if (CFG.SALLES.indexOf(salle) < 0) throw new Error('Salle inconnue.');
  var plan = majPlan_(false), cal = paramsCalendrier_(), today = aujourdhui_();
  var reg = lireRegistre_(salle, force);
  var regs = reglagesActifs_(salle);
  var res = resumeSalle_(salle, reg, plan, cal, today, regs, lireOnglet_('Rappels'), tachesClimat_());
  res.plan = { version: plan.version, nom: plan.nom };
  res.offsets = { jour: cal.offJour, nuit: cal.offNuit };
  res.rounds = reg.rounds.map(function (r) { return { onglet: r.onglet, start: r.start, recolte: r.recolte }; });
  // Round affiché (par défaut : en cours, sinon le plus récent)
  var r = null;
  if (ongletVoulu) r = reg.rounds.filter(function (x) { return x.onglet === ongletVoulu; })[0] || null;
  if (!r && reg.rounds.length) {
    r = reg.rounds.filter(function (x) { return x.start <= today; }).pop() || reg.rounds[reg.rounds.length - 1];
  }
  res.ongletAffiche = r ? r.onglet : null;
  res.editable = !!(r && r.gacp);
  if (r) {
    res.semaines = semainesRound_(r, plan, cal, today);
    res.roundAffiche = { onglet: r.onglet, start: r.start, veg: r.veg, flo: r.flo, recolte: r.recolte, floJ1: plusJours_(r.start, r.veg), fin: r.fin };
  } else if (reg.continu) {
    var cv = cible_(null, 'Veg 1', plan);
    res.semaines = [{ semaine: 'Veg 1', courante: true, cible: cv, attendu: attendu_(cv, cal), source: cv.source }];
  }
  if (r) {
    res.changements = changements_(r, plan).map(function (e) {
      e.confirme = regs.some(function (x) { return x['Round'] == r.round && x['Semaine'] === e.semaine; });
      return e;
    });
    // Comparaison jour par jour jusqu'à aujourd'hui
    var fin = r.fin < today ? r.fin : today, jours = [];
    for (var iso = r.start; iso <= fin; iso = plusJours_(iso, 1)) {
      var j = ecartJours_(r.start, iso) + 1, s = semaineDuJour_(r.veg, r.flo, j);
      if (!s) break;
      var c = cible_(r, s.label, plan), att = attendu_(c, cal), rg = reglageA_(regs, iso), mes = r.mesures[iso] || null;
      jours.push({ date: iso, jour: j, semaine: s.label, cible: { tj: c.tj, hrj: c.hrj, tn: c.tn, hrn: c.hrn, co2: c.co2, hrnMax: c.hrnMax },
        attendu: att, regle: regVal_(rg), mesure: mes, diag: diagnostic_(c, att, rg, mes) });
    }
    res.comparaison = jours.reverse();
    var cpt = {}; jours.forEach(function (x) { cpt[x.diag.code] = (cpt[x.diag.code] || 0) + 1; });
    res.bilan = cpt;
  } else if (reg.continu) {
    var cc = cible_(null, 'Veg 1', plan), at = attendu_(cc, cal);
    var jr = [];
    for (var k = 0; k < 28; k++) {
      var d = plusJours_(today, -k), rg2 = reglageA_(regs, d);
      jr.push({ date: d, semaine: 'Veg 1', cible: cc, attendu: at, regle: regVal_(rg2), mesure: null, diag: diagnostic_(cc, at, rg2, null) });
    }
    res.comparaison = jr;
  }
  res.historique = regs.slice(-60).reverse().map(resumeReglage_);
  var filtre = function (n) { return lireOnglet_(n).filter(function (x) { return x['Salle'] === salle; }).slice(-30).reverse(); };
  res.ir = filtre('Relevés IR');
  res.transitions = filtre('Transition OFF');
  res.observations = filtre('Observations');
  res.annules = lireOnglet_('Réglages').filter(function (x) { return x['Salle'] === salle && x['Statut'] === 'annulé'; }).slice(-20).reverse()
    .map(function (x) { return { id: x['ID'], date: String(x['Date effective']).slice(0, 10), semaine: x['Semaine'], motif: x['Motif annulation'], le: x['Annulé le'] }; });
  res.rapports = rapports_(salle);
  return res;
}

function rapports_(salle) {
  var c = CacheService.getScriptCache(), k = 'rap:' + salle, hit = c.get(k);
  if (hit) return JSON.parse(hit);
  var out = [];
  try {
    var n = salle.replace('Salle ', '');
    var it = DriveApp.getFolderById(CFG.RAPPORTS_FOLDER).searchFiles("title contains 'Salle " + n + "' and trashed = false");
    while (it.hasNext() && out.length < 10) { var f = it.next(); out.push({ nom: f.getName(), url: f.getUrl(), date: f.getLastUpdated().toISOString().slice(0, 10) }); }
    var it2 = DriveApp.getFolderById(CFG.CLIMAT_FOLDER).searchFiles("title contains 'Résumé analyse climat' and trashed = false");
    var res = [];
    while (it2.hasNext()) { var g = it2.next(); res.push({ nom: g.getName(), url: g.getUrl(), date: g.getLastUpdated().toISOString().slice(0, 10) }); }
    res.sort(function (a, b) { return a.nom < b.nom ? 1 : -1; });
    if (res[0]) out.unshift(res[0]);
  } catch (e) {}
  out.sort(function (a, b) { return a.date < b.date ? 1 : -1; });
  c.put(k, JSON.stringify(out), 1800);
  return out;
}

// ───────────────────────────── Saisies ─────────────────────────────
function contexte_(salle, iso) {
  if (CFG.SALLES.indexOf(salle) < 0) throw new Error('Salle inconnue.');
  var plan = majPlan_(false), cal = paramsCalendrier_();
  var reg = lireRegistre_(salle, false);
  var sit = situation_(reg, iso, plan);
  var c = sit.vide ? null : cible_(sit.roundObj || null, sit.semaine, plan);
  return { plan: plan, cal: cal, sit: sit, cible: c, attendu: attendu_(c, cal) };
}

function id_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

function enregistrerReglage_(p) {
  var iso = p.date || aujourdhui_();
  var x = contexte_(p.salle, iso);
  var val = { tj: num_(p.tj), hrj: num_(p.hrj), tn: num_(p.tn), hrn: num_(p.hrn), co2: num_(p.co2), rampe: num_(p.rampe) };
  if ([val.tj, val.hrj, val.tn, val.hrn, val.co2, val.rampe].every(function (n) { return n === null; })) throw new Error('Aucune valeur saisie.');
  if (x.sit.round) { // un plan de rattrapage actif remplace la consigne attendue
    var pa = planActif_(p.salle, x.sit.round), st = pa && pa.etapes.filter(function (e) { return e.date <= iso; }).pop();
    if (st) x.attendu = st.val;
  }
  var verif = verifierReglage_(val, x.attendu);
  var semaine = p.semaine || (x.sit.vide ? '' : x.sit.semaine);
  var c = x.cible || {};
  var id = id_(), now = maintenant_();
  var tache = null;
  if (p.cocherCalendrier) { try { tache = cocherTacheCalendrier_(p.salle, semaine, 'T ' + val.tj + '/' + val.tn + ' °C · HR ' + val.hrj + '/' + val.hrn + ' %'); } catch (e) { tache = { erreur: String(e.message || e) }; } }
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    ss_().getSheetByName('Réglages').appendRow([id, now, iso, p.heure || heure_(), p.salle, x.sit.round || (x.sit.continu ? 'continu' : ''), x.sit.phase || '', semaine,
      val.tj, val.hrj, val.tn, val.hrn, val.co2, val.rampe,
      c.tj, c.hrj, c.tn, c.hrn, c.co2, c.rampe,
      x.attendu ? x.attendu.tj : '', x.attendu ? x.attendu.tn : '', c.source || (x.sit.vide ? 'Salle vide' : ''),
      verif.ok === null ? 'sans cible' : verif.ok ? 'conforme' : 'écart', verif.ecarts.join(' ; '), p.note || '', 'Phil', 'actif', '', '',
      tache && tache.id ? tache.id + ' · ' + tache.titre : '']);
  } finally { lock.releaseLock(); }
  journal_('réglage', p.salle, id, semaine + ' · ' + JSON.stringify(val) + (verif.ecarts.length ? ' · écarts : ' + verif.ecarts.join(' ; ') : ''));
  var registre = null;
  try { registre = ecrireRegleRegistre_(p.salle, iso, val, p.heure || heure_()); } catch (e) { registre = { ok: false, motif: String(e.message || e) }; }
  CacheService.getScriptCache().remove('reg:' + p.salle);
  return { id: id, conformite: verif, tache: tache, registre: registre };
}

function enregistrerIR_(p) {
  var iso = p.date || aujourdhui_();
  var x = contexte_(p.salle, iso);
  var t = num_(p.tAir), hr = num_(p.hr);
  var lect = [num_(p.ir1), num_(p.ir2), num_(p.ir3)].filter(function (n) { return n !== null; });
  if (!lect.length) throw new Error('Inscrivez au moins une lecture IR.');
  var moy = lect.reduce(function (a, b) { return a + b; }, 0) / lect.length;
  var dp = ptRosee_(t, hr);
  var marge = dp === null ? null : moy - dp;
  var seuil = x.plan.params.margeRosee || CFG.TOL.MARGE;
  var verdict = marge === null ? 'T air / HR manquantes' : marge >= seuil ? 'OK' : 'RISQUE condensation';
  var id = id_();
  ss_().getSheetByName('Relevés IR').appendRow([id, maintenant_(), iso, p.heure || heure_(), p.salle, x.sit.round || '', x.sit.semaine || '', num_(p.minutes),
    t, hr, num_(p.ir1), num_(p.ir2), num_(p.ir3), arr1_(moy), t === null ? '' : arr1_(t - moy), arr1_(dp), arr1_(marge), verdict, p.note || '', 'Phil', 'actif', '', '']);
  journal_('relevé IR', p.salle, id, 'T feuille ' + arr1_(moy) + ' °C · ' + verdict);
  return { id: id, tFeuille: arr1_(moy), delta: t === null ? null : arr1_(t - moy), ptRosee: arr1_(dp), marge: arr1_(marge), verdict: verdict };
}

function enregistrerTransition_(p) {
  var iso = p.date || aujourdhui_();
  var x = contexte_(p.salle, iso);
  var g = function (k) { return num_(p[k]); };
  var dp0 = ptRosee_(g('t0'), g('hr0')), dp30 = ptRosee_(g('t30'), g('hr30')), dp90 = ptRosee_(g('t90'), g('hr90'));
  var c30 = dp0 !== null && dp30 !== null ? dp0 - dp30 : null, c90 = dp0 !== null && dp90 !== null ? dp0 - dp90 : null;
  var tr = x.plan.transition[x.sit.semaine] || {}, P = x.plan.params;
  var plafond = tr.plafond, cible = tr.chuteCible;
  var vDep = c30 === null ? '—' : c30 >= (P.baisseDp30 || 0.3) ? 'OK' : 'Lent';
  var v90 = c90 === null || cible === null || cible === undefined ? '—' : c90 >= cible - (P.toleranceDp90 || 0.2) ? 'OK' : 'Insuffisant';
  var rhMax = g('hrMax'), mins = g('minutes');
  var vExp = rhMax === null || plafond === null || plafond === undefined ? '—' : rhMax <= plafond ? 'OK' : 'Dépassé' + (mins !== null ? ' (' + mins + ' min)' : '');
  var dpMax = Math.max.apply(null, [dp0, dp30, dp90].filter(function (n) { return n !== null; }).concat([-99]));
  var surf = g('surface');
  var vCond = surf === null || dpMax === -99 ? '—' : surf - dpMax >= (P.margeRosee || 3) ? 'OK' : 'RISQUE';
  var id = id_();
  ss_().getSheetByName('Transition OFF').appendRow([id, maintenant_(), iso, p.salle, x.sit.round || '', x.sit.semaine || '',
    g('t0'), g('hr0'), g('t30'), g('hr30'), g('t90'), g('hr90'), rhMax, mins, surf,
    arr1_(dp0), arr1_(dp30), arr1_(dp90), arr2_(c30), arr2_(c90), plafond === undefined ? '' : plafond, cible === undefined ? '' : cible,
    vDep, v90, vExp, vCond, p.note || '', 'Phil', 'actif', '', '']);
  journal_('transition OFF', p.salle, id, [vDep, v90, vExp, vCond].join(' / '));
  return { id: id, dp: [arr1_(dp0), arr1_(dp30), arr1_(dp90)], chute30: arr2_(c30), chute90: arr2_(c90), plafond: plafond, chuteCible: cible,
    verdicts: { depart: vDep, a90: v90, exposition: vExp, condensation: vCond } };
}

function enregistrerObservation_(p) {
  if (!p.texte) throw new Error('Observation vide.');
  var iso = p.date || aujourdhui_();
  var x = contexte_(p.salle, iso);
  var id = id_();
  ss_().getSheetByName('Observations').appendRow([id, maintenant_(), iso, p.salle, x.sit.round || '', x.sit.semaine || '', p.categorie || 'Général', p.texte, 'Phil', 'actif', '', '']);
  journal_('observation', p.salle, id, String(p.texte).slice(0, 200));
  return { id: id };
}

/** Annulation (jamais d'effacement : traçabilité GACP). */
function annuler_(p) {
  var onglets = { reglage: 'Réglages', ir: 'Relevés IR', transition: 'Transition OFF', observation: 'Observations' };
  var nom = onglets[p.type];
  if (!nom) throw new Error('Type inconnu.');
  if (!p.motif) throw new Error('Le motif d\'annulation est obligatoire.');
  var sh = ss_().getSheetByName(nom), v = sh.getDataRange().getValues(), h = v[0];
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(p.id)) {
      if (v[i][h.indexOf('Statut')] === 'annulé') throw new Error('Déjà annulé.');
      sh.getRange(i + 1, h.indexOf('Statut') + 1).setValue('annulé');
      sh.getRange(i + 1, h.indexOf('Motif annulation') + 1).setValue(p.motif);
      sh.getRange(i + 1, h.indexOf('Annulé le') + 1).setValue(maintenant_());
      journal_('annulation', nom, p.id, p.motif);
      return { id: p.id };
    }
  }
  throw new Error('Entrée introuvable.');
}

function actionRappel_(p) {
  var id = id_();
  ss_().getSheetByName('Rappels').appendRow([id, maintenant_(), p.salle, p.round, p.semaine, p.date, p.action || 'ignoré', p.note || '']);
  journal_('rappel ' + (p.action || 'ignoré'), p.salle, id, p.semaine + ' · ' + (p.note || ''));
  return { id: id };
}

/** Moyenne des deltas IR par semaine (à reporter dans l'onglet « T feuille nuit » du plan). */
function sommaireIR_() {
  var rows = lireOnglet_('Relevés IR').filter(function (x) { return x['Statut'] !== 'annulé'; });
  var par = {};
  rows.forEach(function (x) {
    var s = x['Semaine'] || '—';
    (par[s] = par[s] || { n: 0, d: 0, f: 0, lectures: [] });
    var d = num_(x['Delta air − feuille (°C)']), f = num_(x['T feuille moy. (°C)']);
    if (d !== null) { par[s].n++; par[s].d += d; }
    if (f !== null) par[s].lectures.push(f);
  });
  return { semaines: SEMAINES.filter(function (s) { return par[s]; }).map(function (s) {
    var o = par[s]; return { semaine: s, n: o.n, deltaMoyen: o.n ? arr1_(o.d / o.n) : null, lectures: o.lectures.slice(-3) };
  }) };
}

// ───────────────────────────── Export GACP ─────────────────────────────
function exporter_(p) {
  var salle = p.salle, onglet = p.onglet || null;
  var regs = lireOnglet_('Réglages').filter(function (x) { return x['Salle'] === salle; });
  var reg = lireRegistre_(salle, false), r = null;
  if (onglet) r = reg.rounds.filter(function (x) { return x.onglet === onglet; })[0] || null;
  if (r) regs = regs.filter(function (x) { return String(x['Round']) === String(r.round); });
  if (!regs.length) throw new Error('Aucun réglage à exporter pour cette sélection.');
  var titre = 'Registre des réglages climat – ' + salle + (r ? ' – Round ' + r.round : ' – ' + aujourdhui_());
  var out = SpreadsheetApp.create(titre);
  var dossier = sousDossierExport_();
  DriveApp.getFileById(out.getId()).moveTo(dossier);
  var sh = out.getSheets()[0];
  sh.setName('Registre');
  var tete = [
    [titre, '', '', '', ''],
    ['Établissement', 'Amaltis Inc.', '', 'Préparé par', 'Philippe Brunelle (chef cultivateur)'],
    ['Salle', salle, '', 'Round', r ? r.round + ' (' + r.onglet + ')' : 'tous'],
    ['Période du round', r ? r.start + ' au ' + r.fin : '—', '', 'Produit le', maintenant_()],
    ['Source des consignes', 'Tableau du registre de culture, sinon Plan_VPD (dernière version)', '', 'Décalage thermostat', 'jour ' + paramsCalendrier_().offJour + ' °C · nuit ' + paramsCalendrier_().offNuit + ' °C'],
    ['Note', 'Les entrées ne sont jamais effacées : une correction est une nouvelle entrée, l\'ancienne est marquée « annulé » avec son motif.', '', '', ''],
    ['', '', '', '', '']
  ];
  sh.getRange(1, 1, tete.length, 5).setValues(tete);
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  sh.getRange(2, 1, 5, 1).setFontWeight('bold'); sh.getRange(2, 4, 4, 1).setFontWeight('bold');
  var cols = ['Date effective', 'Heure', 'Semaine', 'Phase', 'T jour réglée (°C)', 'HR jour réglée (%)', 'T nuit réglée (°C)', 'HR nuit réglée (%)', 'CO₂ réglé (ppm)', 'Rampe réglée (h)',
    'Cible T jour', 'Cible HR jour', 'Cible T nuit', 'Cible HR nuit', 'Cible CO₂', 'Attendu thermostat T jour', 'Attendu thermostat T nuit', 'Source de la cible',
    'Conformité du réglage', 'Écarts', 'Note', 'Saisi par', 'Horodatage', 'Statut', 'Motif annulation', 'Annulé le', 'Tâche calendrier', 'ID'];
  var data = [cols].concat(regs.map(function (x) { return cols.map(function (k) { var v = x[k]; return v === undefined || v === null ? '' : v; }); }));
  var r0 = tete.length + 1;
  sh.getRange(r0, 1, data.length, cols.length).setValues(data);
  sh.getRange(r0, 1, 1, cols.length).setFontWeight('bold').setBackground('#e6ebf0').setWrap(true);
  sh.getRange(r0, 1, data.length, cols.length).setBorder(true, true, true, true, true, true, '#9aa5b1', SpreadsheetApp.BorderStyle.SOLID);
  sh.setFrozenRows(r0);
  var sig = r0 + data.length + 2;
  sh.getRange(sig, 1, 3, 3).setValues([['Signature', 'Nom', 'Date'], ['Chef cultivateur', 'Philippe Brunelle', ''], ['Responsable qualité', '', '']]);
  sh.getRange(sig, 1, 1, 3).setFontWeight('bold');
  journal_('export', salle, out.getId(), titre + ' · ' + regs.length + ' ligne(s)');
  return { url: out.getUrl(), titre: titre, lignes: regs.length };
}

function sousDossierExport_() {
  var parent = DriveApp.getFolderById(CFG.CLIMAT_FOLDER);
  var it = parent.getFoldersByName(CFG.EXPORT_SUBFOLDER);
  return it.hasNext() ? it.next() : parent.createFolder(CFG.EXPORT_SUBFOLDER);
}

// ───────────────────────────── Journal ─────────────────────────────
function journal_(action, objet, id, details) {
  try { ss_().getSheetByName('Journal').appendRow([maintenant_(), action, objet || '', id || '', details || '']); } catch (e) {}
}

// ═════════════════════════════ Écriture dans le registre de culture ═════════════════════════════
var REG_HEAD = ['Réglé T jour (°C)', 'Réglé HR jour (%)', 'Réglé T nuit (°C)', 'Réglé HR nuit (%)', 'Réglé CO₂ (ppm)', 'Réglé rampe (h)', 'Réglé – heure / par'];
var CONSIGNE_HEAD = ['T jour consigne', 'HR jour consigne', 'VPD min consigne', 'VPD max consigne', 'T nuit consigne', 'HR nuit consigne', 'HR nuit max consigne', 'CO₂ consigne (appliquée)'];
var TABLE_OFF = { tj: 2, hrj: 3, tn: 4, hrn: 5, rampe: 10, co2: 11 };
var CHAMPS = { tj: 'T jour', hrj: 'HR jour', tn: 'T nuit', hrn: 'HR nuit', co2: 'CO₂', rampe: 'Rampe' };

function tab_(nom, entetes) {
  var s = ss_(), sh = s.getSheetByName(nom);
  if (!sh) { sh = s.insertSheet(nom); sh.getRange(1, 1, 1, entetes.length).setValues([entetes]).setFontWeight('bold').setBackground('#e6ebf0'); sh.setFrozenRows(1); }
  return sh;
}

/** Ouvre l'onglet du round en cours (ou nommé) d'un registre, avec repères de colonnes. */
function ctxRound_(salle, onglet) {
  if (salle === CFG.SALLE_MERES) return null;
  var book = SpreadsheetApp.openById(paramsCalendrier_().registres[salle]);
  var tz = book.getSpreadsheetTimeZone();
  var tabs = book.getSheets().filter(function (sh) { return /round/i.test(sh.getName()); }).map(function (sh) {
    var b = sh.getRange('B1:B3').getDisplayValues();
    return { sh: sh, start: dateCellule_(b[0][0], tz), veg: num_(b[1][0]), flo: num_(b[2][0]) };
  }).filter(function (t) { return t.start && t.veg !== null && t.flo !== null; }).sort(function (a, b) { return a.start < b.start ? -1 : 1; });
  var today = aujourdhui_(), t = null;
  if (onglet) t = tabs.filter(function (x) { return x.sh.getName() === onglet; })[0];
  if (!t) t = tabs.filter(function (x) { return x.start <= today; }).pop() || tabs[tabs.length - 1];
  if (!t) throw new Error('Aucun round trouvé dans le registre de ' + salle + '.');
  var sh = t.sh, lastCol = sh.getLastColumn(), lastRow = Math.max(7, sh.getLastRow());
  var v = sh.getRange(1, 1, Math.min(lastRow, 220), lastCol).getValues();
  var H = v[5].map(function (h) { return String(h).replace(/\s+/g, ' ').trim(); });
  var ctx = { book: book, sh: sh, onglet: sh.getName(), round: (sh.getName().match(/(\d+)/) || [, sh.getName()])[1], start: t.start, veg: t.veg, flo: t.flo, v: v, H: H, lastCol: lastCol };
  ctx.cDay = H.indexOf('Day');
  ctx.cObs = H.indexOf('Observations / écarts');
  ctx.cCons = CONSIGNE_HEAD.map(function (n) { return H.indexOf(n); });
  ctx.tableCol = -1;
  for (var j = 20; j < lastCol; j++) if (String(v[4][j]).trim() === 'Semaine' && String(v[4][j + 1]).trim() === 'Phase') { ctx.tableCol = j; break; }
  ctx.cReg = H.indexOf(REG_HEAD[0]);
  return ctx;
}

/** Ligne (1-based) du registre pour une date, ou -1 si hors du round. */
function ligneDuJour_(ctx, iso) {
  var j = ecartJours_(ctx.start, iso) + 1;
  if (j < 1 || ctx.cDay < 0) return -1;
  for (var r = 6; r < ctx.v.length; r++) if (num_(ctx.v[r][ctx.cDay]) === j) return r + 1;
  return -1;
}

/** Fige les consignes des jours passés (formules → valeurs) pour que l'historique ne soit pas réécrit. */
function figerPasses_(ctx) {
  var cols = ctx.cCons.filter(function (c) { return c >= 0; });
  if (!cols.length) return 0;
  var hier = plusJours_(aujourdhui_(), -1), fin = ligneDuJour_(ctx, hier);
  if (fin < 0) { // le round a peut-être commencé aujourd'hui, ou il est terminé
    var j = ecartJours_(ctx.start, hier) + 1;
    if (j < 1) return 0;
    fin = 0; for (var r = 6; r < ctx.v.length; r++) if (num_(ctx.v[r][ctx.cDay]) !== null) fin = r + 1;
  }
  if (fin < 7) return 0;
  var c0 = Math.min.apply(null, cols), c1 = Math.max.apply(null, cols);
  var rg = ctx.sh.getRange(7, c0 + 1, fin - 6, c1 - c0 + 1);
  var f = rg.getFormulas(), vals = rg.getValues(), n = 0;
  f.forEach(function (row) { row.forEach(function (x) { if (x) n++; }); });
  if (n) rg.setValues(vals);
  return n;
}

/** Crée au besoin le bloc « Réglé automate » à droite du registre ; renvoie sa 1re colonne (1-based). */
function blocRegle_(ctx) {
  if (ctx.cReg >= 0) return ctx.cReg + 1;
  var last = 0;
  ctx.v[5].forEach(function (x, i) { if (String(x).trim() !== '') last = i; });
  if (ctx.tableCol >= 0) last = Math.max(last, ctx.tableCol + 12);
  var col = last + 3; // une colonne vide de séparation
  var need = col + REG_HEAD.length - 1;
  if (ctx.sh.getMaxColumns() < need) ctx.sh.insertColumnsAfter(ctx.sh.getMaxColumns(), need - ctx.sh.getMaxColumns());
  ctx.sh.getRange(5, col).setValue('Réglé automate — saisi dans l\'app Climat (valeurs entrées au panneau)').setFontWeight('bold');
  ctx.sh.getRange(6, col, 1, REG_HEAD.length).setValues([REG_HEAD]).setFontWeight('bold').setBackground('#dcecf2').setWrap(true);
  ctx.cReg = col - 1;
  return col;
}

function ajouterObservation_(ctx, iso, texte) {
  if (ctx.cObs < 0) return false;
  var row = ligneDuJour_(ctx, iso);
  if (row < 0) return false;
  var cell = ctx.sh.getRange(row, ctx.cObs + 1), cur = String(cell.getValue() || '').trim();
  cell.setValue(cur ? cur + ' · ' + texte : texte);
  return true;
}

function ecrireRegleRegistre_(salle, iso, val, heure) {
  var ctx = ctxRound_(salle);
  if (!ctx) return { ok: false, motif: 'Salle 2 : pas de registre par round' };
  var row = ligneDuJour_(ctx, iso);
  if (row < 0) return { ok: false, motif: 'date hors du round ' + ctx.onglet };
  var col = blocRegle_(ctx);
  var ex = String(ctx.sh.getRange(row, col + 6).getValue() || '');
  var trace = heure + ' · Phil' + (ex ? ' (remplace ' + ex + ')' : '');
  ctx.sh.getRange(row, col, 1, 7).setValues([[val.tj, val.hrj, val.tn, val.hrn, val.co2, val.rampe, trace]]
    .map(function (r) { return r.map(function (x) { return x === null || x === undefined ? '' : x; }); }));
  return { ok: true, onglet: ctx.onglet, ligne: row };
}

// ───────────── Modifier les cibles (tableau de consignes du registre) ─────────────
function modifierCible_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var ctx = ctxRound_(p.salle, p.onglet);
    if (!ctx || ctx.tableCol < 0) throw new Error('Le registre de ' + p.salle + ' n\'a pas de tableau de consignes (format GACP au prochain round).');
    var semaines = (p.semaines || []).filter(function (s) { return SEMAINES.indexOf(s) >= 0; });
    if (!semaines.length) throw new Error('Aucune semaine choisie.');
    var champs = Object.keys(p.valeurs || {}).filter(function (k) { return TABLE_OFF[k] !== undefined && num_(p.valeurs[k]) !== null; });
    if (!champs.length) throw new Error('Aucune valeur à modifier.');
    var figes = figerPasses_(ctx);
    var modifs = [], j = ctx.tableCol;
    semaines.forEach(function (s) {
      for (var r = 4; r < Math.min(ctx.v.length, 30); r++) {
        if (String(ctx.v[r][j]).trim() !== s) continue;
        champs.forEach(function (k) {
          var anc = num_(ctx.v[r][j + TABLE_OFF[k]]), nouv = num_(p.valeurs[k]);
          if (p.mode === 'decaler') nouv = anc === null ? null : arr1_(anc + num_(p.valeurs[k]));
          if (nouv === null || anc === nouv) return;
          ctx.sh.getRange(r + 1, j + TABLE_OFF[k] + 1).setValue(nouv);
          modifs.push({ semaine: s, champ: k, de: anc, a: nouv });
        });
      }
    });
    if (!modifs.length) return { modifs: [], figes: figes };
    var today = aujourdhui_(), now = maintenant_();
    var txt = 'Cible modifiée (app Climat) : ' + resumeModifs_(modifs) + (p.motif ? ' — ' + p.motif : '');
    ajouterObservation_(ctx, today, txt);
    var sh = tab_('Modifs cibles', ['ID', 'Horodatage', 'Salle', 'Round', 'Semaine', 'Champ', 'Ancienne valeur', 'Nouvelle valeur', 'Motif', 'Par']);
    var id = id_();
    sh.getRange(sh.getLastRow() + 1, 1, modifs.length, 10).setValues(modifs.map(function (m) { return [id, now, p.salle, ctx.round, m.semaine, CHAMPS[m.champ], m.de, m.a, p.motif || '', 'Phil']; }));
    journal_('cible', p.salle, id, txt + (figes ? ' · ' + figes + ' consigne(s) passée(s) figée(s)' : ''));
    CacheService.getScriptCache().remove('reg:' + p.salle);
    return { modifs: modifs, figes: figes, id: id };
  } finally { lock.releaseLock(); }
}

function resumeModifs_(modifs) {
  var par = {};
  modifs.forEach(function (m) { var k = CHAMPS[m.champ] + ' ' + m.de + '→' + m.a; (par[k] = par[k] || []).push(m.semaine); });
  return Object.keys(par).map(function (k) { return k + ' (' + par[k].join(', ') + ')'; }).join(' ; ');
}

// ───────────── Assistant de descente (récolte avancée ou reportée) ─────────────
function cloneRound_(r, flo) {
  var o = JSON.parse(JSON.stringify(r));
  o.flo = flo; o.fin = plusJours_(r.start, r.veg + flo - 1); o.recolte = plusJours_(r.start, r.veg + flo);
  return o;
}

/** Plan de rattrapage, jours ouvrables seulement : chaque jour, on se rapproche de la consigne
 *  attendue d'au plus maxT °C et maxHR points d'HR ; CO₂ et rampe suivent directement. */
function rattrapage_(r2, plan, cal, depuis, depart, maxT, maxHR) {
  var etapes = [], cur = JSON.parse(JSON.stringify(depart));
  var vers = function (a, b, m) { if (a === null || a === undefined) return b; if (b === null || b === undefined) return a; var d = b - a; return Math.abs(d) <= m ? b : arr1_(a + (d > 0 ? m : -m)); };
  for (var iso = depuis; iso <= r2.fin; iso = plusJours_(iso, 1)) {
    var dw = jourSemaine_(iso);
    if (dw === 0 || dw === 6) continue;
    var j = ecartJours_(r2.start, iso) + 1, s = semaineDuJour_(r2.veg, r2.flo, j);
    if (!s) break;
    var att = attendu_(cible_(r2, s.label, plan), cal);
    var nxt = { tj: vers(cur.tj, att.tj, maxT), tn: vers(cur.tn, att.tn, maxT), hrj: vers(cur.hrj, att.hrj, maxHR), hrn: vers(cur.hrn, att.hrn, maxHR), co2: att.co2, rampe: att.rampe };
    var change = ['tj', 'tn', 'hrj', 'hrn', 'co2', 'rampe'].some(function (k) { return nxt[k] !== cur[k]; });
    var rejoint = ['tj', 'tn', 'hrj', 'hrn'].every(function (k) { return nxt[k] === att[k]; });
    if (change) etapes.push({ date: iso, semaine: s.label, val: nxt, attendu: att, rejoint: rejoint });
    cur = nxt;
    if (rejoint) break; // consigne rattrapée : la suite du calendrier normal prend le relais
  }
  return etapes;
}

function descente_(p, appliquer) {
  var plan = majPlan_(false), cal = paramsCalendrier_(), today = aujourdhui_();
  var reg = lireRegistre_(p.salle, true);
  var r = reg.rounds.filter(function (x) { return x.start <= today; }).pop();
  if (!r) throw new Error('Aucun round en cours pour ' + p.salle + '.');
  var flo = num_(p.flo);
  if (flo === null || flo < 42 || flo > 77) throw new Error('Nombre de jours de FLO invalide (42 à 77).');
  var r2 = cloneRound_(r, flo);
  var jour = ecartJours_(r.start, today) + 1;
  if (jour > r.veg + flo) throw new Error('Avec ' + flo + ' j de FLO, la récolte serait déjà passée.');
  var s1 = semaineDuJour_(r.veg, r.flo, jour), s2 = semaineDuJour_(r2.veg, r2.flo, jour);
  var regs = reglagesActifs_(p.salle), last = regs.length ? regVal_(regs[regs.length - 1]) : null;
  var depart = last || attendu_(cible_(r, s1 ? s1.label : 'Veg 1', plan), cal);
  var maxT = num_(p.maxT) || 1, maxHR = num_(p.maxHR) || 3;
  var ancien = changements_(r, plan).filter(function (e) { return !e.sansChangement; });
  var nouveau = changements_(r2, plan).filter(function (e) { return !e.sansChangement; });
  var etapes = rattrapage_(r2, plan, cal, today, depart, maxT, maxHR);
  var out = { salle: p.salle, onglet: r.onglet, ancienFlo: r.flo, flo: flo, ancienneRecolte: r.recolte, recolte: r2.recolte,
    semaineAvant: s1 && s1.label, semaineApres: s2 && s2.label, depart: depart, departSource: last ? 'dernier réglage inscrit' : 'consigne attendue avant le changement',
    ancien: ancien, nouveau: nouveau, etapes: etapes, maxT: maxT, maxHR: maxHR };
  if (!appliquer) return out;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var ctx = ctxRound_(p.salle, r.onglet);
    var figes = figerPasses_(ctx);
    ctx.sh.getRange('B3').setValue(flo);
    var txt = 'FLO ' + r.flo + '→' + flo + ' j (récolte ' + r.recolte + ' → ' + r2.recolte + ') ; descente ajustée, ' + etapes.length + ' étape(s) de réglage (app Climat)' + (p.motif ? ' — ' + p.motif : '');
    ajouterObservation_(ctx, today, txt);
    var sh = tab_('Plans', ['ID', 'Horodatage', 'Salle', 'Round', 'Ancien FLO', 'Nouveau FLO', 'Récolte', 'Étapes (JSON)', 'Statut', 'Motif']);
    var id = id_();
    sh.appendRow([id, maintenant_(), p.salle, r.round, r.flo, flo, r2.recolte, JSON.stringify(etapes), 'actif', p.motif || '']);
    journal_('descente', p.salle, id, txt + (figes ? ' · ' + figes + ' consigne(s) passée(s) figée(s)' : ''));
    var c = CacheService.getScriptCache(); c.remove('reg:' + p.salle); c.remove('tachesClimat');
    out.applique = true; out.id = id;
    return out;
  } finally { lock.releaseLock(); }
}

/** Plan de rattrapage actif d'une salle (dernier plan « actif » du round en cours). */
function planActif_(salle, round) {
  var s = ss_().getSheetByName('Plans');
  if (!s || s.getLastRow() < 2) return null;
  var v = s.getRange(2, 1, s.getLastRow() - 1, 10).getValues();
  for (var i = v.length - 1; i >= 0; i--) {
    if (v[i][2] === salle && String(v[i][3]) === String(round) && v[i][8] === 'actif') {
      var et = []; try { et = JSON.parse(v[i][7]); } catch (e) {}
      if (!et.length || et[et.length - 1].date < plusJours_(aujourdhui_(), -3)) return null;
      return { id: v[i][0], le: v[i][1] instanceof Date ? Utilities.formatDate(v[i][1], CFG.TZ, 'yyyy-MM-dd HH:mm') : String(v[i][1]), ancienFlo: v[i][4], flo: v[i][5], etapes: et };
    }
  }
  return null;
}

/** Historique d'une salle : réglages, modifications de cibles, descentes. */
function historique_(salle) {
  var out = [];
  reglagesActifs_(salle).slice(-80).forEach(function (x) { out.push({ type: 'réglage', quand: x['Date effective'] + ' ' + (x['Heure'] || ''), semaine: x['Semaine'], val: regVal_(x), conf: x['Conformité du réglage'], ecarts: x['Écarts'], note: x['Note'], id: x['ID'] }); });
  var mc = ss_().getSheetByName('Modifs cibles');
  if (mc && mc.getLastRow() > 1) mc.getRange(2, 1, mc.getLastRow() - 1, 9).getValues().forEach(function (r) {
    if (r[2] !== salle) return;
    out.push({ type: 'cible', quand: r[1] instanceof Date ? Utilities.formatDate(r[1], CFG.TZ, 'yyyy-MM-dd HH:mm') : String(r[1]), semaine: r[4], champ: r[5], de: r[6], a: r[7], motif: r[8] });
  });
  var pl = ss_().getSheetByName('Plans');
  if (pl && pl.getLastRow() > 1) pl.getRange(2, 1, pl.getLastRow() - 1, 10).getValues().forEach(function (r) {
    if (r[2] !== salle) return;
    out.push({ type: 'descente', quand: r[1] instanceof Date ? Utilities.formatDate(r[1], CFG.TZ, 'yyyy-MM-dd HH:mm') : String(r[1]), de: r[4], a: r[5], recolte: r[6] instanceof Date ? Utilities.formatDate(r[6], CFG.TZ, 'yyyy-MM-dd') : String(r[6]), motif: r[9] });
  });
  out.sort(function (a, b) { return a.quand < b.quand ? 1 : -1; });
  return { items: out.slice(0, 120) };
}

/** Tableau des 12 semaines d'un round avec dates et source. */
function semainesRound_(r, plan, cal, today) {
  var ch = changements_(r, plan), out = [];
  SEMAINES.forEach(function (s) {
    var e = ch.filter(function (x) { return x.semaine === s; })[0];
    var c = cible_(r, s, plan);
    var idx = ch.indexOf(e), fin = e ? (ch[idx + 1] ? plusJours_(ch[idx + 1].date, -1) : r.fin) : null;
    out.push({ semaine: s, debut: e ? e.date : null, conseillee: e ? e.conseillee : null, fin: fin, absente: !e,
      courante: !!(e && e.date <= today && fin >= today), cible: { tj: c.tj, hrj: c.hrj, tn: c.tn, hrn: c.hrn, co2: c.co2, rampe: c.rampe, vpd: c.vpd, vpdMin: c.vpdMin, vpdMax: c.vpdMax, hrnMax: c.hrnMax, dif: c.dif },
      attendu: attendu_(c, cal), divergences: c.divergences, source: c.source });
  });
  return out;
}
