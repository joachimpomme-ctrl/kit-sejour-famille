/**
 * SÉJOUR — serveur Apps Script (kit générique, duplicable)
 * Données stockées en JSON dans un Google Sheet créé automatiquement
 * dans le Drive du propriétaire du script.
 * Toutes les écritures passent par un verrou : pas d'écrasement croisé.
 */

var SHEET_PROP = 'SEJOUR_SHEET_ID';
var FOLDER_PROP = 'SEJOUR_FOLDER_ID';
var ADMINS_PROP = 'SEJOUR_ADMINS';
var ALERTE_PROP = 'SEJOUR_ALERTES';
var CATS_ALERTES_ = ['repas', 'depenses', 'famille', 'chambres', 'wifi', 'courses', 'reglages'];
var CAT_OP_ = {
  setMeal: 'repas',
  addExpense: 'depenses', delExpense: 'depenses', setExpense: 'depenses',
  addParticipant: 'famille', delParticipant: 'famille', setParticipant: 'famille',
  setPresence: 'famille',
  assignRoom: 'chambres',
  setWifi: 'wifi',
  addCourse: 'courses', checkCourse: 'courses', delCourse: 'courses',
  setConfig: 'reglages',
  addBatiment: 'reglages', setBatiment: 'reglages', delBatiment: 'reglages',
  addRoom: 'reglages', setRoom: 'reglages', delRoom: 'reglages',
  addFoyer: 'reglages', renameFoyer: 'reglages', delFoyer: 'reglages'
};

/* ------------------------- Point d'entrée web ------------------------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Séjour')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ------------------------- API appelée par le client ------------------ */

function getData() {
  return JSON.stringify(readData_());
}

function applyOp(opJson) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  var desc, titre, res;
  try {
    var op = JSON.parse(opJson);
    var d = readData_();
    desc = decrireOp_(d, op);
    appliquer_(d, op);
    writeData_(d);
    titre = (d.config && d.config.titre) || 'Séjour';
    res = JSON.stringify(d);
  } finally {
    lock.releaseLock();
  }
  try { notifier_(desc, CAT_OP_[op.t], titre); } catch (e) {}
  return res;
}

/** Reçoit une facture (photo ou PDF) en base64, la dépose dans le dossier
 *  Drive partagé et renvoie l'id du fichier.
 *  Appelée AVANT addExpense ; si elle échoue, la dépense s'enregistre sans PJ. */
function uploadFacture(nom, type, b64) {
  var blob = Utilities.newBlob(
    Utilities.base64Decode(b64),
    String(type || 'application/octet-stream'),
    String(nom || 'facture').slice(0, 90)
  );
  return dossierFactures_().createFile(blob).getId();
}

function dossierFactures_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(FOLDER_PROP);
  var folder = null;
  if (id) {
    try { folder = DriveApp.getFolderById(id); } catch (e) { folder = null; }
  }
  if (!folder) {
    folder = DriveApp.createFolder(nomInstance_() + ' — factures');
    try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    props.setProperty(FOLDER_PROP, folder.getId());
  }
  return folder;
}

/** À lancer UNE FOIS dans l'éditeur après la copie du projet : autorise les
 *  accès (Sheets, Drive, Mail), crée le classeur de données et le dossier
 *  factures. À refaire avant toute mise à jour du déploiement si un nouveau
 *  scope OAuth apparaît, sinon l'app renvoie une page d'autorisation. */
function initSejour() {
  readData_();
  Logger.log('Classeur de données : ' + SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty(SHEET_PROP)).getUrl());
  Logger.log('Dossier factures : ' + dossierFactures_().getUrl());
  Logger.log('Mails d\'alerte envoyés à : ' + Session.getEffectiveUser().getEmail());
}

function nomInstance_() {
  try {
    var d = readData_();
    if (d.config && d.config.titre) return d.config.titre;
  } catch (e) {}
  return 'Séjour';
}

/* ------------------------- Admin (alertes mail) ----------------------- */

function admins_() {
  var p = PropertiesService.getScriptProperties().getProperty(ADMINS_PROP);
  if (p) { try { var l = JSON.parse(p); if (l && l.length) return l; } catch (e) {} }
  return [{ email: Session.getEffectiveUser().getEmail(), pass: 'vacances' }];
}

function verifAdmin_(email, pass) {
  var e = String(email || '').trim().toLowerCase();
  return admins_().some(function (a) {
    return a.email.toLowerCase() === e && a.pass === String(pass || '');
  });
}

function alertesCfg_() {
  var v = PropertiesService.getScriptProperties().getProperty(ALERTE_PROP);
  var cats = {};
  CATS_ALERTES_.forEach(function (k) { cats[k] = 1; });
  if (v === 'off') return { on: false, cats: cats };
  if (v && v !== 'on') {
    try {
      var c = JSON.parse(v);
      if (c && c.cats) {
        CATS_ALERTES_.forEach(function (k) { if (!(k in c.cats)) c.cats[k] = 1; });
        return { on: !!c.on, cats: c.cats };
      }
    } catch (e) {}
  }
  return { on: true, cats: cats };
}

function adminLogin(email, pass) {
  if (!verifAdmin_(email, pass)) return JSON.stringify({ ok: false });
  return JSON.stringify({ ok: true, cfg: alertesCfg_() });
}

function adminSetAlertes(email, pass, cfgJson) {
  if (!verifAdmin_(email, pass)) return JSON.stringify({ ok: false });
  var c = JSON.parse(cfgJson);
  var cfg = { on: !!c.on, cats: {} };
  CATS_ALERTES_.forEach(function (k) { cfg.cats[k] = c.cats && c.cats[k] ? 1 : 0; });
  PropertiesService.getScriptProperties().setProperty(ALERTE_PROP, JSON.stringify(cfg));
  return JSON.stringify({ ok: true, cfg: cfg });
}

function adminSetPass(email, pass, nouveau) {
  if (!verifAdmin_(email, pass)) return JSON.stringify({ ok: false });
  var np = String(nouveau || '').trim();
  if (np.length < 4) return JSON.stringify({ ok: false, err: 'court' });
  var cible = String(email).trim().toLowerCase();
  var list = admins_().map(function (a) {
    return a.email.toLowerCase() === cible ? { email: a.email, pass: np } : a;
  });
  PropertiesService.getScriptProperties().setProperty(ADMINS_PROP, JSON.stringify(list));
  return JSON.stringify({ ok: true });
}

/* ------------------------- Stockage ----------------------------------- */

function cellule_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(SHEET_PROP);
  var ss = null;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('Nouveau séjour — données');
    props.setProperty(SHEET_PROP, ss.getId());
  }
  var sh = ss.getSheetByName('data') || ss.insertSheet('data');
  return sh.getRange('A1');
}

function readData_() {
  var cell = cellule_();
  var v = cell.getValue();
  if (!v) {
    var init = donneesInitiales_();
    cell.setValue(JSON.stringify(init));
    return init;
  }
  try {
    var d = JSON.parse(v);
    if (!d || !d.participants || !d.config) throw new Error('vide');
    return d;
  } catch (e) {
    var init2 = donneesInitiales_();
    cell.setValue(JSON.stringify(init2));
    return init2;
  }
}

function writeData_(d) {
  cellule_().setValue(JSON.stringify(d));
}

/** Renomme le classeur de données et le dossier factures d'après le titre,
 *  pour s'y retrouver dans le Drive quand on a plusieurs séjours. */
function renommerFichiers_(titre) {
  var props = PropertiesService.getScriptProperties();
  try {
    var sid = props.getProperty(SHEET_PROP);
    if (sid) SpreadsheetApp.openById(sid).rename(titre + ' — données');
  } catch (e) {}
  try {
    var fid = props.getProperty(FOLDER_PROP);
    if (fid) DriveApp.getFolderById(fid).setName(titre + ' — factures');
  } catch (e) {}
}

/* ------------------------- Opérations --------------------------------- */

function idCourt_(prefixe) {
  return prefixe + Date.now() + Math.random().toString(36).slice(2, 5);
}

function dateSure_(v, sinon) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(v || '')) ? String(v) : (sinon || '');
}

function appliquer_(d, op) {
  switch (op.t) {
    case 'setMeal': {
      var m = d.meals[op.slot] || { team: '', duo: null, menu: '' };
      if (op.patch.hasOwnProperty('team')) m.team = op.patch.team;
      if (op.patch.hasOwnProperty('duo')) m.duo = op.patch.duo;
      if (op.patch.hasOwnProperty('menu')) m.menu = op.patch.menu;
      if (op.patch.hasOwnProperty('absents')) {
        m.absents = Array.isArray(op.patch.absents)
          ? op.patch.absents.slice(0, 100).map(String) : [];
      }
      d.meals[op.slot] = m;
      break;
    }
    case 'addExpense': {
      var exp = {
        id: String(Date.now()) + '-' + Math.random().toString(36).slice(2, 7),
        ts: Date.now(),
        label: String(op.exp.label || '').slice(0, 120),
        amount: Math.round((parseFloat(op.exp.amount) || 0) * 100) / 100,
        payer: String(op.exp.payer || ''),
        link: String(op.exp.link || 'autre')
      };
      if (op.exp.pj) exp.pj = String(op.exp.pj).slice(0, 80);
      if (Array.isArray(op.exp.pour) && op.exp.pour.length) {
        exp.pour = op.exp.pour.slice(0, 12).map(function (f) { return String(f).slice(0, 40); });
      }
      d.expenses.push(exp);
      break;
    }
    case 'delExpense': {
      d.expenses = d.expenses.filter(function (e) { return e.id !== op.id; });
      break;
    }
    case 'setExpense': {
      d.expenses.forEach(function (e) {
        if (e.id !== op.id) return;
        if (op.hasOwnProperty('link')) e.link = String(op.link || 'autre');
        if (Array.isArray(op.pour) && op.pour.length) {
          e.pour = op.pour.slice(0, 12).map(function (f) { return String(f).slice(0, 40); });
        } else if (op.hasOwnProperty('pour')) delete e.pour;
      });
      break;
    }
    case 'addParticipant': {
      var foyer = String(op.foyer || '').slice(0, 40);
      if (foyer && d.foyers.indexOf(foyer) === -1) d.foyers.push(foyer);
      d.participants.push({
        id: idCourt_('p'),
        nom: String(op.nom || '').slice(0, 40),
        foyer: foyer,
        poids: parseFloat(op.poids) || 0,
        apero: op.hasOwnProperty('apero') ? (op.apero ? 1 : 0) : ((parseFloat(op.poids) || 0) >= 1 ? 1 : 0),
        arrivee: dateSure_(op.arrivee, ''),
        depart: dateSure_(op.depart, '')
      });
      break;
    }
    case 'delParticipant': {
      d.participants = d.participants.filter(function (p) { return p.id !== op.id; });
      if (d.chambres) delete d.chambres[op.id];
      break;
    }
    case 'setParticipant': {
      for (var i = 0; i < d.participants.length; i++) {
        if (d.participants[i].id === op.id) {
          if (op.champ === 'poids') d.participants[i].poids = parseFloat(op.valeur) || 0;
          else if (op.champ === 'foyer') {
            var f = String(op.valeur).slice(0, 40);
            d.participants[i].foyer = f;
            if (d.foyers.indexOf(f) === -1) d.foyers.push(f);
          } else if (op.champ === 'nom') d.participants[i].nom = String(op.valeur).slice(0, 40);
          else if (op.champ === 'apero') d.participants[i].apero = op.valeur && op.valeur !== '0' ? 1 : 0;
          else if (op.champ === 'arrivee') d.participants[i].arrivee = dateSure_(op.valeur, '');
          else if (op.champ === 'depart') d.participants[i].depart = dateSure_(op.valeur, '');
        }
      }
      break;
    }
    case 'setPresence': {
      var ids = op.ids || [];
      d.participants.forEach(function (p) {
        if (ids.indexOf(p.id) === -1) return;
        if (op.patch.hasOwnProperty('arrivee')) p.arrivee = dateSure_(op.patch.arrivee, '');
        if (op.patch.hasOwnProperty('depart')) p.depart = dateSure_(op.patch.depart, '');
      });
      break;
    }
    case 'assignRoom': {
      if (!d.chambres) d.chambres = {};
      if (op.room) d.chambres[op.pid] = op.room;
      else delete d.chambres[op.pid];
      break;
    }
    case 'setWifi': {
      if (!d.wifi) d.wifi = {};
      var w = d.wifi[op.id] || {};
      if (op.patch.hasOwnProperty('ssid')) w.ssid = String(op.patch.ssid || '').slice(0, 60);
      if (op.patch.hasOwnProperty('pass')) w.pass = String(op.patch.pass || '').slice(0, 80);
      d.wifi[op.id] = w;
      break;
    }
    case 'addCourse': {
      if (!d.liste) d.liste = [];
      d.liste.push({ id: idCourt_('c'), txt: String(op.txt || '').slice(0, 80), fait: 0, ts: Date.now() });
      break;
    }
    case 'checkCourse': {
      (d.liste || []).forEach(function (x) { if (x.id === op.id) x.fait = op.fait ? 1 : 0; });
      break;
    }
    case 'delCourse': {
      d.liste = (d.liste || []).filter(function (x) { return x.id !== op.id; });
      break;
    }
    case 'setConfig': {
      var c = op.patch || {};
      if (c.hasOwnProperty('titre')) {
        var t = String(c.titre || '').slice(0, 60).trim();
        if (t) { d.config.titre = t; renommerFichiers_(t); }
      }
      if (c.hasOwnProperty('sousTitre')) d.config.sousTitre = String(c.sousTitre || '').slice(0, 80);
      if (c.hasOwnProperty('debut')) d.config.debut = dateSure_(c.debut, d.config.debut);
      if (c.hasOwnProperty('fin')) d.config.fin = dateSure_(c.fin, d.config.fin);
      if (c.hasOwnProperty('prorata')) d.config.prorata = c.prorata ? 1 : 0;
      break;
    }
    case 'addBatiment': {
      d.batiments.push({ id: idCourt_('b'), nom: String(op.nom || 'Maison').slice(0, 40), rooms: [] });
      break;
    }
    case 'setBatiment': {
      d.batiments.forEach(function (b) {
        if (b.id === op.id && op.patch.hasOwnProperty('nom')) {
          var n = String(op.patch.nom || '').slice(0, 40).trim();
          if (n) b.nom = n;
        }
      });
      break;
    }
    case 'delBatiment': {
      var bat = null;
      d.batiments = d.batiments.filter(function (b) { if (b.id === op.id) { bat = b; return false; } return true; });
      if (bat) {
        bat.rooms.forEach(function (r) { libererChambre_(d, r.id); });
        if (d.wifi) delete d.wifi[op.id];
      }
      break;
    }
    case 'addRoom': {
      d.batiments.forEach(function (b) {
        if (b.id === op.bat) b.rooms.push({ id: idCourt_('r'), nom: String(op.nom || 'Chambre').slice(0, 40), lits: '', cap: 2 });
      });
      break;
    }
    case 'setRoom': {
      d.batiments.forEach(function (b) {
        b.rooms.forEach(function (r) {
          if (r.id !== op.id) return;
          if (op.patch.hasOwnProperty('nom')) {
            var n = String(op.patch.nom || '').slice(0, 40).trim();
            if (n) r.nom = n;
          }
          if (op.patch.hasOwnProperty('lits')) r.lits = String(op.patch.lits || '').slice(0, 60);
          if (op.patch.hasOwnProperty('cap')) r.cap = Math.max(1, Math.min(12, parseInt(op.patch.cap, 10) || 2));
        });
      });
      break;
    }
    case 'delRoom': {
      d.batiments.forEach(function (b) {
        b.rooms = b.rooms.filter(function (r) { return r.id !== op.id; });
      });
      libererChambre_(d, op.id);
      break;
    }
    case 'addFoyer': {
      var nf = String(op.nom || '').slice(0, 40).trim();
      if (nf && d.foyers.indexOf(nf) === -1) d.foyers.push(nf);
      break;
    }
    case 'renameFoyer': {
      var de = String(op.de || ''), vers = String(op.vers || '').slice(0, 40).trim();
      if (!vers || de === vers) break;
      d.foyers = d.foyers.map(function (f) { return f === de ? vers : f; })
        .filter(function (f, i, a) { return a.indexOf(f) === i; });
      d.participants.forEach(function (p) { if (p.foyer === de) p.foyer = vers; });
      d.expenses.forEach(function (e) {
        if (e.payer === de) e.payer = vers;
        if (Array.isArray(e.pour)) e.pour = e.pour.map(function (f) { return f === de ? vers : f; });
      });
      Object.keys(d.meals).forEach(function (k) { if (d.meals[k].team === de) d.meals[k].team = vers; });
      break;
    }
    case 'delFoyer': {
      var occupe = d.participants.some(function (p) { return p.foyer === op.nom; });
      if (!occupe) d.foyers = d.foyers.filter(function (f) { return f !== op.nom; });
      break;
    }
  }
}

function libererChambre_(d, roomId) {
  if (!d.chambres) return;
  Object.keys(d.chambres).forEach(function (pid) {
    if (d.chambres[pid] === roomId) delete d.chambres[pid];
  });
}

/* ------------------------- Alertes mail ------------------------------- */

var JOURS_FR_ = ['dimanche', 'lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi'];
var MOIS_FR_ = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
var SERVICE_LBL_ = { dej: 'déjeuner', din: 'dîner' };

function jourFr_(iso) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(iso || ''))) return String(iso || '');
  var dt = new Date(iso + 'T12:00:00');
  var n = dt.getDate();
  return JOURS_FR_[dt.getDay()] + ' ' + (n === 1 ? '1er' : n) + ' ' + MOIS_FR_[dt.getMonth()];
}

function slotLbl_(slot) {
  var s = String(slot || '');
  var i = s.lastIndexOf('-');
  if (i === -1) return s;
  return (SERVICE_LBL_[s.slice(i + 1)] || s.slice(i + 1)) + ' du ' + jourFr_(s.slice(0, i));
}

function nomPar_(d, id) {
  for (var i = 0; i < d.participants.length; i++) {
    if (d.participants[i].id === id) return d.participants[i].nom;
  }
  return null;
}

function nomBat_(d, id) {
  for (var i = 0; i < d.batiments.length; i++) {
    if (d.batiments[i].id === id) return d.batiments[i].nom;
  }
  return null;
}

function nomChambre_(d, roomId) {
  for (var i = 0; i < d.batiments.length; i++) {
    for (var j = 0; j < d.batiments[i].rooms.length; j++) {
      if (d.batiments[i].rooms[j].id === roomId) {
        return d.batiments[i].nom + ' · ' + d.batiments[i].rooms[j].nom;
      }
    }
  }
  return roomId;
}

/** Décrit l'opération en français, à partir des données AVANT application
 *  (nécessaire pour nommer ce qui va être supprimé). */
function decrireOp_(d, op) {
  switch (op.t) {
    case 'setMeal':
      if (op.patch && op.patch.hasOwnProperty('absents')) {
        var abs = (op.patch.absents || []).map(function (id) { return nomPar_(d, id); }).filter(Boolean);
        return 'Absents du ' + slotLbl_(op.slot) + ' : ' + (abs.length ? abs.join(', ') : 'plus personne');
      }
      if (op.patch && op.patch.hasOwnProperty('menu')) {
        var menu = String(op.patch.menu || '').slice(0, 80);
        return 'Menu du ' + slotLbl_(op.slot) + ' : ' + (menu || '(effacé)');
      }
      if (op.patch && Array.isArray(op.patch.duo)) {
        var noms = op.patch.duo.map(function (id) { return nomPar_(d, id); }).filter(Boolean);
        return 'Cuisine du ' + slotLbl_(op.slot) + ' : ' + (noms.length ? noms.join(' + ') : 'duo à choisir');
      }
      if (op.patch && op.patch.hasOwnProperty('team')) {
        return 'Cuisine du ' + slotLbl_(op.slot) + ' : ' + (op.patch.team || 'à pourvoir');
      }
      return 'Repas modifié : ' + slotLbl_(op.slot);
    case 'addExpense':
      return 'Dépense ajoutée : ' + String(op.exp && op.exp.label || '').slice(0, 80) +
        ', ' + (parseFloat(op.exp && op.exp.amount) || 0) + ' € (' + String(op.exp && op.exp.payer || '?') + ')' +
        (op.exp && Array.isArray(op.exp.pour) && op.exp.pour.length ? ', pour ' + op.exp.pour.join(' + ') : '') +
        (op.exp && op.exp.pj ? ', avec facture' : '');
    case 'delExpense': {
      var e = null;
      d.expenses.forEach(function (x) { if (x.id === op.id) e = x; });
      return e ? 'Dépense supprimée : ' + e.label + ', ' + e.amount + ' € (' + e.payer + ')' : 'Dépense supprimée';
    }
    case 'setExpense': {
      var e2 = null;
      d.expenses.forEach(function (x) { if (x.id === op.id) e2 = x; });
      var vers2 = Array.isArray(op.pour) && op.pour.length ? 'pour ' + op.pour.join(' + ')
        : ({ courses: 'courses communes', apero: 'apéro/alcool', autre: 'autre' })[op.link] || slotLbl_(op.link);
      return 'Dépense réaffectée : ' + (e2 ? e2.label : op.id) + ' → ' + vers2;
    }
    case 'addParticipant':
      return 'Participant ajouté : ' + String(op.nom || '').slice(0, 40) + ' (' + String(op.foyer || '').slice(0, 40) + ')';
    case 'delParticipant':
      return 'Participant retiré : ' + (nomPar_(d, op.id) || op.id);
    case 'setParticipant': {
      var champs = { poids: 'parts', foyer: 'ménage', nom: 'nom', apero: 'apéro/alcool', arrivee: 'arrivée', depart: 'départ' };
      var val = (op.champ === 'arrivee' || op.champ === 'depart') ? (jourFr_(op.valeur) || 'tout le séjour') : String(op.valeur).slice(0, 40);
      return 'Fiche de ' + (nomPar_(d, op.id) || op.id) + ' : ' + (champs[op.champ] || op.champ) + ' → ' + val;
    }
    case 'setPresence': {
      var champP = op.patch && op.patch.hasOwnProperty('arrivee') ? 'arrivee' : 'depart';
      var valP = jourFr_(op.patch && op.patch[champP]) || 'tout le séjour';
      return 'Présence du ménage ' + String(op.foyer || '?').slice(0, 40) + ' : ' +
        (champP === 'arrivee' ? 'arrivée' : 'départ') + ' → ' + valP;
    }
    case 'assignRoom': {
      var nom = nomPar_(d, op.pid) || op.pid;
      return op.room ? nom + ' → ' + nomChambre_(d, op.room) : nom + ' : retiré(e) de sa chambre';
    }
    case 'setWifi':
      return 'Wifi ' + (nomBat_(d, op.id) || op.id) + ' mis à jour';
    case 'addCourse':
      return 'Pense-bête : + ' + String(op.txt || '').slice(0, 80);
    case 'checkCourse': {
      var it = null;
      (d.liste || []).forEach(function (x) { if (x.id === op.id) it = x; });
      return 'Pense-bête : ' + (it ? it.txt : op.id) + (op.fait ? ' acheté' : ' remis à acheter');
    }
    case 'delCourse': {
      var it2 = null;
      (d.liste || []).forEach(function (x) { if (x.id === op.id) it2 = x; });
      return 'Pense-bête : ' + (it2 ? it2.txt : op.id) + ' retiré';
    }
    case 'setConfig': {
      var p = op.patch || {};
      if (p.hasOwnProperty('titre')) return 'Titre du séjour : ' + String(p.titre || '').slice(0, 60);
      if (p.hasOwnProperty('debut') || p.hasOwnProperty('fin')) return 'Dates du séjour modifiées';
      if (p.hasOwnProperty('prorata')) return 'Comptes au prorata : ' + (p.prorata ? 'activé' : 'coupé');
      return 'Réglages du séjour modifiés';
    }
    case 'addBatiment':
      return 'Logement ajouté : ' + String(op.nom || 'Maison').slice(0, 40);
    case 'setBatiment':
      return 'Logement renommé : ' + (nomBat_(d, op.id) || '') + ' → ' + String(op.patch && op.patch.nom || '').slice(0, 40);
    case 'delBatiment':
      return 'Logement supprimé : ' + (nomBat_(d, op.id) || op.id);
    case 'addRoom':
      return 'Chambre ajoutée dans ' + (nomBat_(d, op.bat) || '?');
    case 'setRoom':
      return 'Chambre modifiée : ' + nomChambre_(d, op.id);
    case 'delRoom':
      return 'Chambre supprimée : ' + nomChambre_(d, op.id);
    case 'addFoyer':
      return 'Ménage ajouté : ' + String(op.nom || '').slice(0, 40);
    case 'renameFoyer':
      return 'Ménage renommé : ' + String(op.de || '') + ' → ' + String(op.vers || '').slice(0, 40);
    case 'delFoyer':
      return 'Ménage supprimé : ' + String(op.nom || '');
  }
  return null;
}

function notifier_(desc, cat, titre) {
  if (!desc) return;
  var cfg = alertesCfg_();
  if (!cfg.on) return;
  if (cat && cfg.cats && !cfg.cats[cat]) return;
  if (MailApp.getRemainingDailyQuota() < 5) return;
  var dest = Session.getEffectiveUser().getEmail();
  if (!dest) return;
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  var quand = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM HH:mm');
  MailApp.sendEmail(dest, titre + ' · ' + desc, desc + '\nLe ' + quand + (url ? '\n\n' + url : ''));
}

/* ------------------------- Données initiales -------------------------- */

function donneesInitiales_() {
  var auj = new Date();
  var fin = new Date(auj.getTime() + 7 * 86400000);
  function iso(dt) { return Utilities.formatDate(dt, 'Europe/Paris', 'yyyy-MM-dd'); }
  return {
    version: 5,
    config: {
      titre: 'Notre séjour',
      sousTitre: '',
      debut: iso(auj),
      fin: iso(fin),
      prorata: 0
    },
    batiments: [{ id: 'b1', nom: 'La maison', rooms: [] }],
    foyers: [],
    participants: [],
    meals: {},
    expenses: [],
    chambres: {},
    wifi: {},
    liste: []
  };
}
