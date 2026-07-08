/**
 * ================================================================
 * PATCH SÉCURITÉ — Juillet 2026
 * ================================================================
 *
 * Trois corrections :
 *   1. Verrou anti-doublon sur les numéros de commande (LockService)
 *   2. Téléphone préservé en texte (apostrophe) dans la saisie admin
 *   3. Reçus signés : sans signature, téléphone masqué et adresse cachée
 *
 * 6 SECTIONS — suis-les dans l'ordre. Après : Ctrl+S puis
 * Déployer > Gérer les déploiements > ✏️ > Nouvelle version.
 * ================================================================
 */


/* ================== SECTION 1 — doGet : 1 ligne à modifier ==================
 *
 * Dans le switch de doGet, remplace la ligne :
 *   case 'receipt':            return jsonResponse(getReceipt_(e.parameter.id));
 * par :
 *   case 'receipt':            return jsonResponse(getReceipt_(e.parameter.id, e.parameter.k));
 */


/* ================== SECTION 2 — doPost : 1 ligne à modifier ==================
 *
 * Dans doPost, remplace la ligne :
 *   const orderId = 'JBC-' + (1000 + sheet.getLastRow());
 * par :
 *   const orderId = nextOrderId_(sheet);
 */


/* ================== SECTION 3 — REMPLACE adminCreateOrder_ ================== */

function adminCreateOrder_(params) {
  const nom       = String(params.nom || '').trim();
  const telephone = String(params.telephone || '').trim();
  const ville     = String(params.ville || '').trim();
  const adresse   = String(params.adresse || '').trim();
  const articles  = String(params.articles || '').trim();
  const sousTotal = Number(params.sousTotal) || 0;
  const livraison = Number(params.livraison) || 0;
  const total     = Number(params.total) || 0;
  const paiement  = String(params.paiement || 'Cash').trim();
  const notes     = String(params.notes || '').trim();
  if (!nom || !telephone) return { success:false, error:'Nom et téléphone obligatoires' };
  if (!articles)          return { success:false, error:'Articles obligatoires' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return { success:false, error:'Feuille Commandes introuvable' };
  const orderId = nextOrderId_(sh);
  const now = new Date();
  // Apostrophe devant le téléphone : force le format texte, préserve le 0 initial
  sh.appendRow([now, orderId, nom, "'" + telephone, '', ville, adresse, articles, sousTotal, livraison, total, paiement, 'Nouveau', notes, 'Admin manuel']);
  try {
    const deltas = JSON.parse(params.stockDeltas || '[]');
    if (Array.isArray(deltas) && deltas.length > 0) decrementStock_(deltas);
  } catch (e) {}
  try { sendAdminOrderEmail_({ orderId, nom, telephone, ville, articles, total, paiement, notes }); } catch (e) {}
  try { invalidateAdminCache_(); } catch (e) {}
  return { success:true, id: orderId, date: now.toISOString(), receiptUrl: receiptUrl_(orderId) };
}


/* ================== SECTION 4 — REMPLACE adminOrders_ ================== */

function adminOrders_(filter) {
  const validFilter = ['today', 'pending', 'week', 'all'].indexOf(filter) >= 0 ? filter : 'all';
  return cached_('orders_' + validFilter, function () {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sh) return { success:false, error:'Feuille Commandes introuvable' };
    const rows = toObjects_(sh);
    const today = startOfDay_(new Date());
    const weekStart = daysAgo_(6);
    const filtered = rows.filter(r => {
      const d = parseDate_(r['Date']);
      const statut = String(r['Statut'] || '').toLowerCase();
      if (validFilter === 'today')   return d && d >= today;
      if (validFilter === 'pending') return /nouveau|confirm/.test(statut);
      if (validFilter === 'week')    return d && d >= weekStart;
      return true;
    });
    filtered.sort((a,b) => {
      const da = parseDate_(a['Date']); const db = parseDate_(b['Date']);
      return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
    });
    const orders = filtered.slice(0, 100).map(r => ({
      id: r['ID'] || '',
      date: (parseDate_(r['Date']) || new Date()).toISOString(),
      nom: r['Nom'] || '',
      telephone: r['Telephone'] || '',
      commune: r['Ville'] || '',
      articles: r['Articles'] || '',
      total: parseAmount_(r['Total']),
      statut: r['Statut'] || 'Nouveau',
      receiptUrl: receiptUrl_(r['ID'] || '')
    }));
    return { success:true, orders, total: orders.length };
  });
}


/* ================== SECTION 5 — REMPLACE getReceipt_ ================== */

function getReceipt_(id, k) {
  if (!id) return { success:false, error:'ID manquant' };
  // Accès signé = lien officiel envoyé par Jacob's Bio (WhatsApp, admin).
  // Accès non signé (ID deviné) = coordonnées masquées.
  const signed = !!k && k === receiptSig_(id);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) return { success:false, error:'Feuille Commandes introuvable' };
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idIdx = headers.indexOf('ID');
  if (idIdx === -1) return { success:false, error:'Colonne ID introuvable' };
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idIdx]).trim() === String(id).trim()) {
      const row = values[r];
      const get = name => { const i = headers.indexOf(name); return i >= 0 ? row[i] : ''; };
      const dateVal = get('Date');
      const tel = String(get('Telephone') || '');
      return {
        success: true,
        id: get('ID'),
        date: dateVal instanceof Date ? Utilities.formatDate(dateVal, TIMEZONE, 'dd/MM/yyyy à HH:mm') : String(dateVal),
        nom: get('Nom'),
        telephone: signed ? tel : maskPhone_(tel),
        ville: get('Ville'),
        adresse: signed ? get('Adresse') : '',
        articles: String(get('Articles') || '').split('\n').filter(Boolean),
        sousTotal: Number(String(get('Sous-total')).replace(/[^0-9.-]/g, '')) || 0,
        livraison: Number(String(get('Livraison')).replace(/[^0-9.-]/g, '')) || 0,
        total: Number(String(get('Total')).replace(/[^0-9.-]/g, '')) || 0,
        paiement: get('Paiement'),
        statut: get('Statut') || 'Nouveau'
      };
    }
  }
  return { success:false, error:'Commande introuvable' };
}


/* ================== SECTION 6 — AJOUTE à la fin du fichier ================== */

// ---- Compteur atomique de numéros de commande ----
// LockService + ScriptProperties : deux commandes simultanées (site + admin)
// ne peuvent plus recevoir le même JBC-XXXX.
function nextOrderId_(sheet) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    let counter = parseInt(props.getProperty('ORDER_COUNTER') || '0', 10);
    if (!counter) {
      // Première utilisation : on scanne la feuille pour trouver le max existant
      const values = sheet.getDataRange().getValues();
      const headers = values[0].map(h => String(h).trim());
      const idIdx = headers.indexOf('ID');
      counter = 1000;
      for (let r = 1; r < values.length; r++) {
        const m = String(values[r][idIdx] || '').match(/JBC-(\d+)/);
        if (m) { const n = parseInt(m[1], 10); if (n > counter) counter = n; }
      }
    }
    counter += 1;
    props.setProperty('ORDER_COUNTER', String(counter));
    return 'JBC-' + counter;
  } finally {
    lock.releaseLock();
  }
}

// ---- Signature des liens de reçu ----
function receiptSig_(id) {
  const secret = PropertiesService.getScriptProperties().getProperty('ADMIN_SECRET') || 'default';
  const sig = Utilities.computeHmacSha256Signature('receipt:' + String(id), secret)
    .map(b => (b & 0xff).toString(16).padStart(2, '0')).join('');
  return sig.substring(0, 12);
}

function receiptUrl_(id) {
  return 'https://jacobs-bio.vercel.app/recu.html?id=' + encodeURIComponent(id) + '&k=' + receiptSig_(id);
}

function maskPhone_(t) {
  const s = String(t == null ? '' : t).trim();
  if (!s) return '';
  if (s.length <= 2) return '••';
  return '•'.repeat(s.length - 2) + s.slice(-2);
}


/* ================== SECTION 7 (BONUS) — Liens WhatsApp de reçus ==================
 *
 * Ctrl+F dans Code.gs : cherche  recu.html
 * Partout où un lien est construit à la main, du genre :
 *   'https://jacobs-bio.vercel.app/recu.html?id=' + order.id
 * remplace par :
 *   receiptUrl_(order.id)
 * Ainsi les reçus envoyés par WhatsApp porteront la signature et
 * afficheront les coordonnées complètes de la cliente.
 */
