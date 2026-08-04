// ════════════════════════════════════════════════════════════
//  API USSS v3 — Registre + Pointeuse
//  À coller dans Extensions → Apps Script puis :
//  Déployer → Gérer les déploiements → ✎ → Nouvelle version
// ════════════════════════════════════════════════════════════
//
//  Onglet UTILISATEURS (créé au 1er accès) :
//    Col A : identifiant éditeur     Col B : mot de passe éditeur
//    Col C : matricule agent (E-XX)  Col D : mot de passe agent
//
//  Lecture : libre.
//  Écriture du registre : éditeurs seulement.
//  Pointage : agent (uniquement pour son matricule) ou éditeur.
//  Jetons valables 6 h.
// ════════════════════════════════════════════════════════════

var SHEET_DATA = "STOCKAGE_SITE";
var SHEET_USERS = "UTILISATEURS";

function getDataSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_DATA);
  if (!sh) sh = ss.insertSheet(SHEET_DATA);
  return sh;
}

function getUsersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_USERS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_USERS);
    sh.getRange("A1:D1")
      .setValues([["editeur_id","editeur_mdp","agent_matricule","agent_mdp"]])
      .setFontWeight("bold");
    sh.getRange("A2:B2").setValues([["admin","usss2026"]]);
  } else if (sh.getLastColumn() < 4) {
    // Migration : ajouter les colonnes C/D si onglet préexistant
    sh.getRange("C1:D1")
      .setValues([["agent_matricule","agent_mdp"]])
      .setFontWeight("bold");
  }
  return sh;
}

function checkEditorLogin_(login, password) {
  var sh = getUsersSheet_();
  var last = sh.getLastRow();
  if (last < 2) return false;
  var rows = sh.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    var l = String(rows[i][0]).trim();
    var p = String(rows[i][1]);
    if (l !== "" && l === String(login).trim() && p === String(password)) return true;
  }
  return false;
}

function checkAgentLogin_(matricule, password) {
  var sh = getUsersSheet_();
  var last = sh.getLastRow();
  if (last < 2) return false;
  var rows = sh.getRange(2, 3, last - 1, 2).getValues();
  var target = normaliseMatricule_(matricule);
  for (var i = 0; i < rows.length; i++) {
    var m = normaliseMatricule_(rows[i][0]);
    var p = String(rows[i][1]);
    if (m !== "" && m === target && p === String(password)) return true;
  }
  return false;
}

// Toujours stocker au format E-NN strict : "e16", "E 16", "E-016" → "E-16"
function normaliseMatricule_(s) {
  if (!s) return "";
  var clean = String(s).trim().toUpperCase().replace(/\s+/g, "");
  var m = clean.match(/^E-?(\d{1,3})$/);
  if (!m) return clean;
  return "E-" + ("00" + parseInt(m[1], 10)).slice(-2);
}

// Cellules limitées à 50 000 caractères : JSON découpé en colonne
function writeChunks_(sh, col, str) {
  // Ne nettoyer que les lignes réellement occupées : effacer getMaxRows() lignes
  // (1000 par défaut) à chaque écriture allonge inutilement la tenue du verrou.
  var occupees = Math.max(sh.getLastRow(), 1);
  sh.getRange(1, col, occupees, 1).clearContent();
  var size = 45000, chunks = [];
  for (var i = 0; i < str.length; i += size) chunks.push([str.substring(i, i + size)]);
  // insertRows(1, n) insérerait AVANT la ligne 1 et décalerait toutes les autres
  // colonnes : il faut ajouter les lignes à la fin.
  if (chunks.length > sh.getMaxRows()) {
    sh.insertRowsAfter(sh.getMaxRows(), chunks.length - sh.getMaxRows());
  }
  if (chunks.length) sh.getRange(1, col, chunks.length, 1).setValues(chunks);
}

function readChunks_(sh, col) {
  var last = Math.max(sh.getLastRow(), 1);
  var vals = sh.getRange(1, col, last, 1).getValues();
  var out = "";
  for (var i = 0; i < vals.length; i++) out += vals[i][0];
  return out;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET : lecture libre ──
//   col 1 = registre, col 2 = journal, col 3 = pointages, col 4 = rapports
// Toute exception non rattrapée ferait renvoyer une page HTML d'erreur par Google,
// que le front ne peut pas parser → on garantit une réponse JSON dans tous les cas.
function doGet(e) {
  try {
    return handleGet_(e);
  } catch (err) {
    return out_({ error: "Erreur serveur (doGet) : " + messageErreur_(err) });
  }
}

function messageErreur_(err) {
  return (err && err.message) ? err.message : String(err);
}

// Un seul verrou global protège toutes les écritures. Les blobs sont volumineux :
// 10 s ne suffisent pas quand deux enregistrements s'enchaînent. On attend plus
// longtemps, et surtout on relâche TOUJOURS le verrou, même en cas d'exception.
var LOCK_TIMEOUT_MS = 30000;

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT_MS);
  } catch (err) {
    return out_({ error: "Le registre est occupé par une autre sauvegarde — réessaie dans quelques secondes" });
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function handleGet_(e) {
  var sh = getDataSheet_();
  // Création paresseuse de l'onglet UTILISATEURS : c'est une écriture, elle échoue si
  // le déploiement s'exécute au nom du visiteur. La lecture ne doit pas en dépendre.
  try { getUsersSheet_(); } catch (err) {}
  var members = null, logs = [], pointages = [], rapports = [], autorites = [], divroles = {};
  try { var s = readChunks_(sh, 1); if (s) members = JSON.parse(s); } catch (err) {}
  try { var s2 = readChunks_(sh, 2); if (s2) logs = JSON.parse(s2); } catch (err) {}
  try { var s3 = readChunks_(sh, 3); if (s3) pointages = JSON.parse(s3); } catch (err) {}
  try { var s4 = readChunks_(sh, 4); if (s4) rapports = JSON.parse(s4); } catch (err) {}
  try { var s5 = readChunks_(sh, 5); if (s5) autorites = JSON.parse(s5); } catch (err) {}
  try { var s6 = readChunks_(sh, 6); if (s6) divroles = JSON.parse(s6); } catch (err) {}
  return out_({ members: members, logs: logs, pointages: pointages, rapports: rapports, autorites: autorites, divroles: divroles });
}

function genereNumeroRapport_(rapports) {
  var year = new Date().getFullYear();
  var prefix = "RAP-" + year + "-";
  var max = 0;
  for (var i = 0; i < rapports.length; i++) {
    if (rapports[i].numero && rapports[i].numero.indexOf(prefix) === 0) {
      var n = parseInt(rapports[i].numero.substring(prefix.length), 10);
      if (n > max) max = n;
    }
  }
  return prefix + ("00" + (max + 1)).slice(-3);
}

function loadRapports_(sh) {
  var s = readChunks_(sh, 4);
  try { return s ? JSON.parse(s) : []; } catch (err) { return []; }
}

// ── POST : login & écritures ──
function doPost(e) {
  try {
    return handlePost_(e);
  } catch (err) {
    return out_({ error: "Erreur serveur (doPost) : " + messageErreur_(err) });
  }
}

function handlePost_(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return out_({ error: "JSON invalide" }); }

  var cache = CacheService.getScriptCache();

  // Connexion éditeur
  if (body.action === "login") {
    if (!checkEditorLogin_(body.login, body.password))
      return out_({ error: "Identifiant ou mot de passe incorrect" });
    var token = Utilities.getUuid();
    cache.put("sess_" + token, String(body.login).trim(), 21600);
    return out_({ token: token, login: String(body.login).trim() });
  }

  // Connexion agent
  if (body.action === "login_agent") {
    if (!checkAgentLogin_(body.matricule, body.password))
      return out_({ error: "Matricule ou mot de passe incorrect" });
    var tokenA = Utilities.getUuid();
    var mat = normaliseMatricule_(body.matricule);
    cache.put("agt_" + tokenA, mat, 21600);
    return out_({ token: tokenA, matricule: mat });
  }

  // Sauvegarde du registre (éditeur uniquement)
  if (body.action === "save") {
    var who = cache.get("sess_" + body.session);
    if (!who) return out_({ error: "Session expirée — reconnecte-toi" });
    return withLock_(function() {
      var sh = getDataSheet_();
      writeChunks_(sh, 1, JSON.stringify(body.members));
      writeChunks_(sh, 2, JSON.stringify(body.logs));
      return out_({ ok: true, user: who });
    });
  }

  // Sauvegarde des autorités protégées (éditeur uniquement)
  if (body.action === "save_autorites") {
    var whoA = cache.get("sess_" + body.session);
    if (!whoA) return out_({ error: "Session expirée — reconnecte-toi" });
    return withLock_(function() {
      var shA = getDataSheet_();
      writeChunks_(shA, 5, JSON.stringify(body.autorites || []));
      return out_({ ok: true });
    });
  }

  // Sauvegarde de la gestion des divisions (Lead agent OU éditeur)
  if (body.action === "save_division") {
    var agentMatV = cache.get("agt_" + body.session);
    var editorWhoV = cache.get("sess_" + body.session);
    if (!agentMatV && !editorWhoV) return out_({ error: "Session expirée — reconnecte-toi" });
    return withLock_(function() {
      var shV = getDataSheet_();
      if (body.members) writeChunks_(shV, 1, JSON.stringify(body.members));
      if (body.divroles) writeChunks_(shV, 6, JSON.stringify(body.divroles));
      return out_({ ok: true });
    });
  }

  // Ajout d'un pointage (agent pour lui-même, éditeur pour n'importe qui)
  if (body.action === "add_pointage") {
    var agentMat = cache.get("agt_" + body.session);
    var editorWho = cache.get("sess_" + body.session);
    if (!agentMat && !editorWho) return out_({ error: "Session expirée — reconnecte-toi" });
    var p = body.pointage;
    if (!p || !p.matricule) return out_({ error: "Pointage invalide" });
    p.matricule = normaliseMatricule_(p.matricule);
    if (agentMat && p.matricule !== agentMat)
      return out_({ error: "Tu ne peux pointer que pour ton propre matricule" });

    return withLock_(function() {
      var sh1 = getDataSheet_();
      var s1 = readChunks_(sh1, 3);
      var arr = [];
      try { if (s1) arr = JSON.parse(s1); } catch (err) {}
      arr.unshift(p);
      if (arr.length > 500) arr = arr.slice(0, 500); // garde-fou
      writeChunks_(sh1, 3, JSON.stringify(arr));
      return out_({ ok: true });
    });
  }

  // Suppression d'un pointage
  if (body.action === "delete_pointage") {
    var agentMat2 = cache.get("agt_" + body.session);
    var editorWho2 = cache.get("sess_" + body.session);
    if (!agentMat2 && !editorWho2) return out_({ error: "Session expirée" });

    return withLock_(function() {
      var sh2 = getDataSheet_();
      var s2b = readChunks_(sh2, 3);
      var arr2 = [];
      try { if (s2b) arr2 = JSON.parse(s2b); } catch (err) {}
      var filtered = arr2.filter(function(p) {
        if (p.id !== body.pointageId) return true;
        // Agent : ne peut supprimer que les siens
        if (agentMat2 && p.matricule !== agentMat2) return true;
        return false;
      });
      writeChunks_(sh2, 3, JSON.stringify(filtered));
      return out_({ ok: true });
    });
  }

  // ─── Rapports ───
  // Sauvegarde d'un brouillon (création ou édition par le rédacteur)
  if (body.action === "save_rapport") {
    var agentMatR = cache.get("agt_" + body.session);
    var editorWhoR = cache.get("sess_" + body.session);
    if (!agentMatR && !editorWhoR) return out_({ error: "Session expirée — reconnecte-toi" });
    var r = body.rapport;
    if (!r || !r.id) return out_({ error: "Rapport invalide" });
    // Un agent ne peut sauvegarder que SES propres rapports en brouillon
    if (agentMatR) {
      r.redacteurMat = agentMatR;
      if (r.etat && r.etat !== "brouillon") {
        return out_({ error: "Un rapport soumis n'est plus modifiable" });
      }
    }
    r.etat = r.etat || "brouillon";

    return withLock_(function() {
      var shR = getDataSheet_();
      var rapports = loadRapports_(shR);
      // Si rapport existant : refuser modification si déjà soumis (sauf éditeur)
      var idx = -1;
      for (var i = 0; i < rapports.length; i++) {
        if (rapports[i].id === r.id) { idx = i; break; }
      }
      if (idx >= 0) {
        if (rapports[idx].etat === "soumis" && agentMatR) {
          return out_({ error: "Ce rapport est déjà soumis — verrouillé" });
        }
        if (agentMatR && rapports[idx].redacteurMat !== agentMatR) {
          return out_({ error: "Tu ne peux modifier que tes propres rapports" });
        }
        rapports[idx] = r;
      } else {
        rapports.unshift(r);
        if (rapports.length > 500) rapports = rapports.slice(0, 500);
      }
      writeChunks_(shR, 4, JSON.stringify(rapports));
      return out_({ ok: true });
    });
  }

  // Soumission d'un rapport (verrouille + génère le numéro)
  if (body.action === "submit_rapport") {
    var agentMatS = cache.get("agt_" + body.session);
    var editorWhoS = cache.get("sess_" + body.session);
    if (!agentMatS && !editorWhoS) return out_({ error: "Session expirée — reconnecte-toi" });
    if (!body.rapportId) return out_({ error: "Identifiant manquant" });

    return withLock_(function() {
      var shS = getDataSheet_();
      var rapportsS = loadRapports_(shS);
      var idxS = -1;
      for (var j = 0; j < rapportsS.length; j++) {
        if (rapportsS[j].id === body.rapportId) { idxS = j; break; }
      }
      if (idxS < 0) return out_({ error: "Rapport introuvable" });
      var target = rapportsS[idxS];
      if (target.etat === "soumis") return out_({ error: "Déjà soumis" });
      if (agentMatS && target.redacteurMat !== agentMatS) {
        return out_({ error: "Tu ne peux soumettre que tes propres rapports" });
      }
      target.etat = "soumis";
      target.numero = genereNumeroRapport_(rapportsS);
      target.dateSoumission = Utilities.formatDate(new Date(),
        Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
      rapportsS[idxS] = target;
      writeChunks_(shS, 4, JSON.stringify(rapportsS));
      return out_({ ok: true, numero: target.numero, dateSoumission: target.dateSoumission });
    });
  }

  // Suppression d'un rapport
  //   Agent : peut supprimer SES brouillons uniquement
  //   Éditeur : peut tout supprimer
  if (body.action === "delete_rapport") {
    var agentMatD = cache.get("agt_" + body.session);
    var editorWhoD = cache.get("sess_" + body.session);
    if (!agentMatD && !editorWhoD) return out_({ error: "Session expirée" });

    return withLock_(function() {
      var shD = getDataSheet_();
      var rapportsD = loadRapports_(shD);
      var filtered = rapportsD.filter(function(r) {
        if (r.id !== body.rapportId) return true;
        if (editorWhoD) return false; // éditeur supprime
        if (agentMatD && r.redacteurMat === agentMatD && r.etat === "brouillon") return false;
        return true; // sinon on garde
      });
      writeChunks_(shD, 4, JSON.stringify(filtered));
      return out_({ ok: true });
    });
  }

  return out_({ error: "Action inconnue" });
}
