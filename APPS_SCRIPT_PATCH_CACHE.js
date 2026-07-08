/**
 * ================================================================
 * PATCH CACHE — Optimisation Phase 2A
 * ================================================================
 *
 * Ce patch AJOUTE une couche de cache 60 secondes sur les fonctions
 * lentes (adminStats_, adminOrders_, adminLowStock_) et invalide le
 * cache automatiquement après un update de commande.
 *
 * Résultat attendu :
 *   - 1er appel  : ~10 secondes (comme avant)
 *   - Appels suivants dans 60s : ~300 ms
 *   - Après update (livrée/annulée) : cache purgé, fresh data
 *
 *
 * COMMENT L'APPLIQUER
 * ================================================================
 *
 * Dans ton Code.gs, tu vas :
 *
 *   1. TROUVER et REMPLACER 3 fonctions existantes :
 *      - adminStats_
 *      - adminOrders_
 *      - adminLowStock_
 *
 *   2. TROUVER et REMPLACER la fonction adminUpdateOrderStatus_
 *      pour ajouter l'invalidation de cache.
 *
 *   3. AJOUTER 2 helpers en dessous : cached_() et invalidateAdminCache_()
 *
 *   4. Sauvegarder (Ctrl+S).
 *
 *   5. Redéployer : Déployer > Gérer les déploiements > ✏️ > Nouvelle version.
 *
 *
 * Astuce : ouvre Ctrl+F et cherche `function adminStats_` pour trouver
 * chaque fonction rapidement.
 * ================================================================
 */


// ================== HELPERS CACHE (À AJOUTER À LA FIN DU FICHIER) ==================

const ADMIN_CACHE_TTL_SECONDS = 60;
const ADMIN_CACHE_PREFIX = 'jb_admin_';

/**
 * Wrapper de cache : si la clé existe → retourne la valeur cachée.
 * Sinon → exécute fn(), stocke le résultat, le retourne.
 */
function cached_(key, fn) {
  const cache = CacheService.getScriptCache();
  const fullKey = ADMIN_CACHE_PREFIX + key;
  const hit = cache.get(fullKey);
  if (hit) {
    try {
      const parsed = JSON.parse(hit);
      parsed._cached = true;
      return parsed;
    } catch (e) { /* cache corrompu, on refait */ }
  }
  const result = fn();
  if (result && result.success) {
    try { cache.put(fullKey, JSON.stringify(result), ADMIN_CACHE_TTL_SECONDS); } catch (e) { /* trop gros, on skip */ }
  }
  return result;
}

/**
 * Purge tout le cache admin (à appeler après un write).
 */
function invalidateAdminCache_() {
  const cache = CacheService.getScriptCache();
  const keys = ['stats', 'orders_today', 'orders_pending', 'orders_week', 'orders_all', 'lowstock']
    .map(k => ADMIN_CACHE_PREFIX + k);
  cache.removeAll(keys);
}


// ================== REMPLACE `adminStats_` ==================

function adminStats_() {
  return cached_('stats', function () {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sh) return { success:false, error:'Feuille Commandes introuvable' };
    const rows = toObjects_(sh);
    const today = startOfDay_(new Date());
    const monthStart = startOfMonth_(new Date());
    const weekStart = daysAgo_(6);
    let caJour=0, caMois=0, commandesJour=0, commandesMois=0, commandesSemaine=0, pending=0;
    rows.forEach(r => {
      const d = parseDate_(r['Date']);
      const total = parseAmount_(r['Total']);
      const statut = String(r['Statut'] || '').toLowerCase();
      const annulee = /annul/.test(statut);
      if (d && d >= today && !annulee) { caJour += total; commandesJour++; }
      if (d && d >= monthStart && !annulee) { caMois += total; commandesMois++; }
      if (d && d >= weekStart && !annulee) commandesSemaine++;
      if (/nouveau|confirm/.test(statut)) pending++;
    });
    const stockSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stock');
    let stockLow = 0;
    if (stockSh) {
      toObjects_(stockSh).forEach(r => {
        const q = parseAmount_(r['Stock'] || r['Quantité'] || r['Qte']);
        if (q <= ADMIN_LOW_STOCK_THRESHOLD) stockLow++;
      });
    }
    return { success:true, caJour, caMois, commandesJour, commandesMois, commandesSemaine, pending, stockLow };
  });
}


// ================== REMPLACE `adminOrders_` ==================

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
      statut: r['Statut'] || 'Nouveau'
    }));
    return { success:true, orders, total: orders.length };
  });
}


// ================== REMPLACE `adminLowStock_` ==================

function adminLowStock_() {
  return cached_('lowstock', function () {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stock');
    if (!sh) return { success:false, error:'Feuille Stock introuvable' };
    const items = toObjects_(sh)
      .map(r => ({ nom: r['Nom produit'] || r['Nom'] || r['Produit'] || '', stock: parseAmount_(r['Stock'] || r['Quantité'] || r['Qte']) }))
      .filter(it => it.nom && it.stock <= ADMIN_LOW_STOCK_THRESHOLD)
      .sort((a, b) => a.stock - b.stock);
    return { success:true, items };
  });
}


// ================== REMPLACE `adminUpdateOrderStatus_` ==================

function adminUpdateOrderStatus_(id, newStatus) {
  if (!id || !newStatus) return { success:false, error:'id ou status manquant' };
  const valid = ['Nouveau','Confirmée','Expédiée','Livrée','Annulée'];
  if (valid.indexOf(newStatus) === -1) return { success:false, error:'Statut invalide' };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) return { success:false, error:'Feuille Commandes introuvable' };
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idIdx = headers.indexOf('ID');
  const statutIdx = headers.indexOf('Statut');
  if (idIdx === -1 || statutIdx === -1) return { success:false, error:'Colonnes ID ou Statut introuvables' };
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idIdx]).trim() === String(id).trim()) {
      sh.getRange(r + 1, statutIdx + 1).setValue(newStatus);
      invalidateAdminCache_(); // 🔑 Purge le cache après un write
      return { success:true, id, status:newStatus };
    }
  }
  return { success:false, error:'Commande introuvable: ' + id };
}
