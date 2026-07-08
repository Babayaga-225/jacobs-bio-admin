/**
 * ================================================================
 * PATCH PHASE 2B — Saisie commande depuis l'admin
 * ================================================================
 *
 * Ce patch ajoute 2 nouveaux endpoints :
 *
 *   - ?action=clients       → liste des clientes uniques pour autocomplete
 *   - ?action=createOrder   → crée une nouvelle commande depuis l'admin
 *
 * COMMENT L'APPLIQUER
 * ================================================================
 *
 *   1. Modifier doGet — ajoute 2 nouveaux `case` dans le switch.
 *   2. Ajouter les 3 fonctions helpers à la fin du fichier.
 *   3. Sauvegarder (Ctrl+S).
 *   4. Redéployer : Déployer > Gérer les déploiements > ✏️ > Nouvelle version.
 * ================================================================
 */


// ================== 1. MODIFICATION DU doGet ==================
// Dans ton switch dans doGet, AJOUTE ces 2 lignes juste avant `default:`
//
//   case 'clients':            return jsonResponse(withAuth_(e, adminClients_));
//   case 'createorder':        return jsonResponse(withAuth_(e, () => adminCreateOrder_(e.parameter)));
//
// Le switch complet ressemblera à ça :
/*
function doGet(e) {
  const action = (e.parameter.action || '').toLowerCase();
  try {
    switch (action) {
      case '':                   return jsonResponse({ success:true, message:"Jacob's Bio CRM API en ligne", time:new Date().toISOString() });
      case 'receipt':            return jsonResponse(getReceipt_(e.parameter.id));
      case 'login':              return jsonResponse(adminLogin_(e.parameter.pwd));
      case 'stats':              return jsonResponse(withAuth_(e, adminStats_));
      case 'orders':             return jsonResponse(withAuth_(e, () => adminOrders_(e.parameter.filter)));
      case 'updateorderstatus':  return jsonResponse(withAuth_(e, () => adminUpdateOrderStatus_(e.parameter.id, e.parameter.status)));
      case 'lowstock':           return jsonResponse(withAuth_(e, adminLowStock_));
      case 'clients':            return jsonResponse(withAuth_(e, adminClients_));                                       // ← NEW
      case 'createorder':        return jsonResponse(withAuth_(e, () => adminCreateOrder_(e.parameter)));                // ← NEW
      default:                   return jsonResponse({ success:false, error:'Action inconnue: ' + action });
    }
  } catch (err) {
    return jsonResponse({ success:false, error: String(err && err.message || err) });
  }
}
*/


// ================== 2. FONCTIONS À AJOUTER À LA FIN ==================


/**
 * Retourne la liste des clientes uniques (nom + téléphone + ville)
 * en dédupliquant sur le téléphone.
 * Cache 5 min car ça change lentement.
 */
function adminClients_() {
  return cached_('clients', function () {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sh) return { success:false, error:'Feuille Commandes introuvable' };
    const rows = toObjects_(sh);
    const seen = new Set();
    const clients = [];
    // Dernières commandes en premier → dernière ville/nom en cas de conflit
    rows.reverse().forEach(r => {
      const tel = String(r['Telephone'] || '').replace(/\s+/g, '').trim();
      const nom = String(r['Nom'] || '').trim();
      if (!nom || !tel) return;
      if (seen.has(tel)) return;
      seen.add(tel);
      clients.push({ nom, telephone: tel, ville: String(r['Ville'] || '').trim() });
    });
    // Trier alphabétiquement
    clients.sort((a, b) => a.nom.localeCompare(b.nom, 'fr'));
    return { success:true, clients, total: clients.length };
  });
}


/**
 * Crée une nouvelle commande depuis l'admin.
 * Génère un ID JBC-XXXX auto, envoie l'email notif, décrémente le stock.
 */
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

  // ID auto — on cherche le dernier JBC-XXXX pour incrémenter
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idIdx = headers.indexOf('ID');
  let maxNum = 1000;
  for (let r = 1; r < values.length; r++) {
    const id = String(values[r][idIdx] || '');
    const m = id.match(/JBC-(\d+)/);
    if (m) { const n = parseInt(m[1], 10); if (n > maxNum) maxNum = n; }
  }
  const orderId = 'JBC-' + (maxNum + 1);
  const now = new Date();

  // Ordre EXACT des colonnes A→O
  const row = [
    now,          // A: Date
    orderId,      // B: ID
    nom,          // C: Nom
    telephone,    // D: Telephone
    '',           // E: Email (vide pour saisie admin)
    ville,        // F: Ville
    adresse,      // G: Adresse
    articles,     // H: Articles
    sousTotal,    // I: Sous-total
    livraison,    // J: Livraison
    total,        // K: Total
    paiement,     // L: Paiement
    'Nouveau',    // M: Statut
    notes,        // N: Notes
    'Admin manuel' // O: Source
  ];
  sh.appendRow(row);

  // Décrément stock si des deltas ont été fournis
  try {
    const deltas = JSON.parse(params.stockDeltas || '[]');
    if (Array.isArray(deltas) && deltas.length > 0) decrementStock_(deltas);
  } catch (e) { /* pas bloquant */ }

  // Email notif à Mme Kouassi
  try { sendAdminOrderEmail_({ orderId, nom, telephone, ville, articles, total, paiement, notes }); }
  catch (e) { /* pas bloquant */ }

  // Purge le cache admin pour que le dashboard reflète la nouvelle commande immédiatement
  try { invalidateAdminCache_(); } catch (e) {}

  return { success:true, id: orderId, date: now.toISOString() };
}


/**
 * Décrémente le stock des produits vendus depuis l'onglet Stock.
 * deltas = [{ produit: 'Crème Vitesse 300g', qty: 2 }, ...]
 */
function decrementStock_(deltas) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stock');
  if (!sh) return;
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const nameIdx = headers.findIndex(h => /(nom.?produit|^nom$|^produit$)/i.test(h));
  const stockIdx = headers.findIndex(h => /^stock$|quantit|qte/i.test(h));
  if (nameIdx === -1 || stockIdx === -1) return;

  const nameToRow = {};
  for (let r = 1; r < values.length; r++) {
    const n = String(values[r][nameIdx] || '').trim().toLowerCase();
    if (n) nameToRow[n] = r + 1; // 1-indexed row
  }

  deltas.forEach(d => {
    const rowNum = nameToRow[String(d.produit || '').trim().toLowerCase()];
    if (!rowNum) return;
    const cell = sh.getRange(rowNum, stockIdx + 1);
    const current = Number(cell.getValue()) || 0;
    cell.setValue(Math.max(0, current - Number(d.qty || 0)));
  });
}


/**
 * Envoie un email de notification pour une nouvelle commande admin.
 */
function sendAdminOrderEmail_(d) {
  const subject = `🛒 Commande admin ${d.orderId} · ${d.nom} · ${Number(d.total).toLocaleString('fr-FR')} FCFA`;
  const htmlBody = `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#F5F2ED;padding:20px;border-radius:8px">
      <h2 style="color:#14100C;margin-top:0">Nouvelle commande (saisie admin)</h2>
      <p><strong>N°</strong> : ${d.orderId}<br>
         <strong>Cliente</strong> : ${d.nom}<br>
         <strong>Téléphone</strong> : ${d.telephone}<br>
         <strong>Commune</strong> : ${d.ville}<br>
         <strong>Total</strong> : ${Number(d.total).toLocaleString('fr-FR')} FCFA<br>
         <strong>Paiement</strong> : ${d.paiement}</p>
      <div style="background:white;padding:14px;border-radius:6px;white-space:pre-line">${d.articles}</div>
      ${d.notes ? `<p style="color:#5a4a3c;font-style:italic;margin-top:12px">Notes : ${d.notes}</p>` : ''}
      <div style="text-align:center;color:#888;font-size:11px;margin-top:20px">
        Saisie via l'espace admin · CRM Jacob's Bio
      </div>
    </div>
  `;
  MailApp.sendEmail({ to: NOTIF_EMAIL, subject, htmlBody });
}
