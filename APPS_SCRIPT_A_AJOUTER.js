/**
 * ================================================================
 * JACOB'S BIO CRM — Extensions Admin Phase 2A
 * ================================================================
 *
 * Copier-coller CE bloc DANS ton Code.gs existant (à la suite),
 * puis :
 *   1. Dans Apps Script > Déployer > Nouvelle version
 *   2. Coche "Toute personne" pour l'accès
 *   3. Copie la nouvelle URL /exec si elle change
 *      (normalement elle reste la même après "Nouvelle version")
 *
 * Ajoute AUSSI dans les Propriétés du script :
 *   Fichier > Paramètres du projet > Propriétés du script > +Ajouter
 *   - Clé : ADMIN_PASSWORD    Valeur : (choisis un mot de passe fort)
 *   - Clé : ADMIN_SECRET      Valeur : (une chaîne aléatoire longue,
 *                                       ex: 47djHK92mLqPz8vXn3B)
 *
 * IMPORTANT — Le doGet(e) existant doit maintenant router selon action.
 * Si tu as déjà un `function doGet(e)` qui gère uniquement 'receipt',
 * remplace-le par la version ci-dessous.
 * ================================================================
 */


// ================== CONFIG ==================
const ADMIN_LOW_STOCK_THRESHOLD = 5;
const ADMIN_SESSION_TTL_HOURS = 24;


// ================== ROUTER PRINCIPAL ==================
// Remplace TON doGet existant par celui-ci.
function doGet(e) {
  const action = (e.parameter.action || '').toLowerCase();
  try {
    switch (action) {
      case 'receipt':            return json_(getReceipt_(e.parameter.id));
      case 'login':              return json_(adminLogin_(e.parameter.pwd));
      case 'stats':              return json_(withAuth_(e, adminStats_));
      case 'orders':             return json_(withAuth_(e, () => adminOrders_(e.parameter.filter)));
      case 'updateorderstatus':  return json_(withAuth_(e, () => adminUpdateOrderStatus_(e.parameter.id, e.parameter.status)));
      case 'lowstock':           return json_(withAuth_(e, adminLowStock_));
      default:                   return json_({ success:false, error:'Action inconnue: ' + action });
    }
  } catch (err) {
    return json_({ success:false, error: String(err && err.message || err) });
  }
}


// ================== AUTH ==================
function adminLogin_(pwd) {
  const props = PropertiesService.getScriptProperties();
  const expected = props.getProperty('ADMIN_PASSWORD');
  if (!expected) return { success:false, error:'ADMIN_PASSWORD non configuré' };
  if (!pwd || pwd !== expected) return { success:false, error:'Mot de passe incorrect' };
  const token = generateToken_();
  return { success:true, token, expiresIn: ADMIN_SESSION_TTL_HOURS * 3600 };
}

function generateToken_() {
  const secret = PropertiesService.getScriptProperties().getProperty('ADMIN_SECRET') || 'default-secret';
  const payload = Date.now() + ':' + Utilities.getUuid();
  const sig = Utilities.computeHmacSha256Signature(payload, secret)
    .map(b => (b & 0xff).toString(16).padStart(2, '0')).join('');
  return Utilities.base64EncodeWebSafe(payload) + '.' + sig.substring(0, 32);
}

function verifyToken_(token) {
  if (!token || typeof token !== 'string') return false;
  const parts = token.split('.');
  if (parts.length !== 2) return false;
  try {
    const decoded = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
    const ts = parseInt(decoded.split(':')[0], 10);
    if (!ts) return false;
    const ageMs = Date.now() - ts;
    if (ageMs > ADMIN_SESSION_TTL_HOURS * 3600 * 1000) return false;
    const secret = PropertiesService.getScriptProperties().getProperty('ADMIN_SECRET') || 'default-secret';
    const sig = Utilities.computeHmacSha256Signature(decoded, secret)
      .map(b => (b & 0xff).toString(16).padStart(2, '0')).join('');
    return parts[1] === sig.substring(0, 32);
  } catch (e) { return false; }
}

function withAuth_(e, fn) {
  const token = e.parameter.token;
  if (!verifyToken_(token)) return { success:false, expired:true, error:'Session expirée' };
  return fn();
}


// ================== HELPERS ==================
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function toObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function startOfDay_(d) { const x = new Date(d); x.setHours(0,0,0,0); return x; }
function startOfMonth_(d) { const x = new Date(d); x.setDate(1); x.setHours(0,0,0,0); return x; }
function daysAgo_(n) { const x = new Date(); x.setDate(x.getDate() - n); x.setHours(0,0,0,0); return x; }

function parseDate_(v) {
  if (v instanceof Date) return v;
  if (!v) return null;
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

function parseAmount_(v) {
  if (typeof v === 'number') return v;
  if (!v) return 0;
  return Number(String(v).replace(/[^0-9.-]/g, '')) || 0;
}


// ================== STATS (KPIs) ==================
function adminStats_() {
  const sh = getSheet_('Commandes');
  if (!sh) return { success:false, error:'Feuille Commandes introuvable' };
  const rows = toObjects_(sh);

  const today = startOfDay_(new Date());
  const monthStart = startOfMonth_(new Date());
  const weekStart = daysAgo_(6);

  let caJour = 0, caMois = 0, commandesJour = 0, commandesMois = 0, commandesSemaine = 0, pending = 0;

  rows.forEach(r => {
    const d = parseDate_(r['Date'] || r['Date commande'] || r['date']);
    const total = parseAmount_(r['Total'] || r['Montant'] || r['total']);
    const statut = String(r['Statut'] || r['statut'] || '').toLowerCase();
    const annulee = /annul/.test(statut);

    if (d && d >= today && !annulee) { caJour += total; commandesJour++; }
    if (d && d >= monthStart && !annulee) { caMois += total; commandesMois++; }
    if (d && d >= weekStart && !annulee) { commandesSemaine++; }
    if (/nouveau|confirm/.test(statut)) pending++;
  });

  // stock bas
  const stockSh = getSheet_('Stock');
  let stockLow = 0;
  if (stockSh) {
    const stockRows = toObjects_(stockSh);
    stockRows.forEach(r => {
      const q = parseAmount_(r['Stock'] || r['Quantité'] || r['Qte']);
      if (q <= ADMIN_LOW_STOCK_THRESHOLD) stockLow++;
    });
  }

  return {
    success: true,
    caJour, caMois,
    commandesJour, commandesMois, commandesSemaine,
    pending, stockLow
  };
}


// ================== ORDERS LIST ==================
function adminOrders_(filter) {
  const sh = getSheet_('Commandes');
  if (!sh) return { success:false, error:'Feuille Commandes introuvable' };
  const rows = toObjects_(sh);

  const today = startOfDay_(new Date());
  const weekStart = daysAgo_(6);

  const filtered = rows.filter(r => {
    const d = parseDate_(r['Date'] || r['Date commande'] || r['date']);
    const statut = String(r['Statut'] || r['statut'] || '').toLowerCase();
    if (filter === 'today')   return d && d >= today;
    if (filter === 'pending') return /nouveau|confirm/.test(statut);
    if (filter === 'week')    return d && d >= weekStart;
    return true; // all
  });

  // Trier du plus récent au plus ancien
  filtered.sort((a, b) => {
    const da = parseDate_(a['Date'] || a['Date commande'] || a['date']);
    const db = parseDate_(b['Date'] || b['Date commande'] || b['date']);
    return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
  });

  const mapped = filtered.slice(0, 100).map(r => ({
    id:        r['N° commande'] || r['ID'] || r['id'] || '',
    date:      (parseDate_(r['Date'] || r['Date commande'] || r['date']) || new Date()).toISOString(),
    nom:       r['Nom'] || r['Client'] || r['nom'] || '',
    telephone: r['Téléphone'] || r['Telephone'] || r['Tél'] || r['tel'] || '',
    commune:   r['Commune'] || r['Ville'] || r['commune'] || '',
    articles:  r['Articles'] || r['Produits'] || r['articles'] || '',
    total:     parseAmount_(r['Total'] || r['Montant'] || r['total']),
    statut:    r['Statut'] || r['statut'] || 'Nouveau'
  }));

  return { success:true, orders: mapped, total: mapped.length };
}


// ================== UPDATE ORDER STATUS ==================
function adminUpdateOrderStatus_(id, newStatus) {
  if (!id || !newStatus) return { success:false, error:'id ou status manquant' };
  const validStatus = ['Nouveau', 'Confirmée', 'Expédiée', 'Livrée', 'Annulée'];
  if (validStatus.indexOf(newStatus) === -1) return { success:false, error:'Statut invalide' };

  const sh = getSheet_('Commandes');
  if (!sh) return { success:false, error:'Feuille Commandes introuvable' };

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idColIdx = headers.findIndex(h => /(N°.?commande|^ID$|^id$)/i.test(h));
  const statusColIdx = headers.findIndex(h => /statut/i.test(h));
  if (idColIdx === -1 || statusColIdx === -1) {
    return { success:false, error:'Colonnes N° commande ou Statut introuvables' };
  }

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idColIdx]).trim() === String(id).trim()) {
      sh.getRange(r + 1, statusColIdx + 1).setValue(newStatus);
      return { success:true, id, status:newStatus };
    }
  }
  return { success:false, error:'Commande introuvable: ' + id };
}


// ================== LOW STOCK ==================
function adminLowStock_() {
  const sh = getSheet_('Stock');
  if (!sh) return { success:false, error:'Feuille Stock introuvable' };
  const rows = toObjects_(sh);
  const items = rows
    .map(r => ({
      nom:   r['Nom produit'] || r['Nom'] || r['Produit'] || r['nom'] || '',
      stock: parseAmount_(r['Stock'] || r['Quantité'] || r['Qte'])
    }))
    .filter(it => it.nom && it.stock <= ADMIN_LOW_STOCK_THRESHOLD)
    .sort((a, b) => a.stock - b.stock);

  return { success:true, items };
}


// ================== CORS pour requêtes cross-origin ==================
// Apps Script ContentService JSON accepte déjà les GET cross-origin sans souci.
// Aucune config supplémentaire nécessaire.
