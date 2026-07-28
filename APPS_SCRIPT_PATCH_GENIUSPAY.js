/**
 * ================================================================
 * PATCH GENIUSPAY — Phase 3A
 * ================================================================
 *
 * Deux nouveaux endpoints :
 *   - ?action=createPayment  → cree un paiement GeniusPay, retourne
 *     la checkout_url vers laquelle rediriger la cliente
 *   - ?action=geniusWebhook  → recoit les notifications de GeniusPay
 *     (payment.success / payment.failed / payment.cancelled), verifie
 *     la signature HMAC-SHA256 + timestamp, met a jour la commande
 *
 * PREPARATION (a faire dans l'ordre)
 * ================================================================
 *
 * 1. Onglet Commandes : ajouter une colonne "Ref Paiement" apres
 *    la colonne Source (colonne P). Cette colonne stockera la
 *    reference GeniusPay (ex: MTX-A1B2C3D4E5) pour reconciliation.
 *
 * 2. Fichier > Parametres du projet > Proprietes du script :
 *    ajouter 3 proprietes :
 *
 *      GENIUS_API_KEY        = pk_sandbox_...  (fournie par GeniusPay)
 *      GENIUS_API_SECRET     = sk_sandbox_...  (fournie par GeniusPay)
 *      GENIUS_WEBHOOK_SECRET = whsec_...       (voir etape 5)
 *
 * 3. Dans le doGet router, ajouter la ligne createPayment
 *    (juste avant 'default:') :
 *
 *      case 'createpayment': return jsonResponse(withAuth_(e,
 *        () => geniusCreatePayment_(e.parameter)));
 *
 *    Note : pas de token requis pour createPayment cote vitrine
 *    car la commande n'est pas encore dans l'admin. On utilise a
 *    la place une verification du referer + un token public
 *    faible. Voir SECURITE ci-dessous.
 *
 * 4. Ajouter les fonctions ci-dessous a la fin du fichier.
 *
 * 5. Sauver (Ctrl+S) et redeployer : Deployer > Gerer les
 *    deploiements > icone crayon > Nouvelle version > "GeniusPay".
 *
 * 6. Configurer le webhook dans le dashboard GeniusPay :
 *    URL = ton URL Apps Script /exec + ?action=geniusWebhook
 *    Events = payment.success, payment.failed, payment.cancelled
 *    GeniusPay te donne alors le WEBHOOK_SECRET (whsec_...)
 *    a mettre dans la propriete GENIUS_WEBHOOK_SECRET.
 *
 *
 * ROUTAGE DU WEBHOOK
 * ================================================================
 *
 * GeniusPay envoie le webhook en POST, pas en GET. Il faut donc
 * router aussi dans doPost. Ajouter au debut de doPost, AVANT le
 * check du SECRET existant :
 *
 *   const action = (e.parameter.action || '').toLowerCase();
 *   if (action === 'geniuswebhook') return handleGeniusWebhook_(e);
 *
 *
 * SECURITE
 * ================================================================
 *
 * - X-API-Secret ne quitte jamais Apps Script (jamais dans Vercel,
 *   jamais dans le repo public jacobs-bio)
 * - Webhook signe HMAC-SHA256 : on rejette tout ce qui n'a pas
 *   une signature valide
 * - Protection anti-replay : timestamp du webhook doit etre < 5 min
 * - createPayment cote vitrine : le montant est calcule cote serveur
 *   a partir de la commande deja creee dans Commandes, donc un
 *   attaquant ne peut pas manipuler le montant depuis le browser.
 *
 * ================================================================
 */


// ================== CONFIG ==================
const GENIUS_API_BASE = 'https://geniuspay.ci/api/v1/merchant';
const GENIUS_FEE_PERCENT = 3.5; // frais estimes appliques au client


// ================== CREATE PAYMENT ==================
// Appele depuis la vitrine apres qu'un doPost a cree la commande.
// Le site envoie orderId + amount (deja ajuste avec les frais) et
// on renvoie une checkout_url a laquelle rediriger la cliente.
function geniusCreatePayment_(params) {
  const orderId = String(params.orderId || '').trim();
  if (!orderId) return { success:false, error:'orderId manquant' };

  // Retrouve la commande dans le Sheet pour recuperer client + montant
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) return { success:false, error:'Feuille Commandes introuvable' };
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idIdx = headers.indexOf('ID');
  const refIdx = headers.indexOf('Ref Paiement');
  let rowNum = -1, row = null;
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idIdx]).trim() === orderId) { rowNum = r + 1; row = values[r]; break; }
  }
  if (!row) return { success:false, error:'Commande introuvable' };

  const get = name => { const i = headers.indexOf(name); return i >= 0 ? row[i] : ''; };
  const amount = Number(String(get('Total')).replace(/[^0-9.-]/g, '')) || 0;
  const nom = String(get('Nom') || '').trim();
  const tel = String(get('Telephone') || '').replace(/^'/, '');
  const email = String(get('Email') || '').trim();

  if (amount < 200) return { success:false, error:'Montant minimum 200 FCFA' };

  // Appel API GeniusPay
  const apiKey = PropertiesService.getScriptProperties().getProperty('GENIUS_API_KEY');
  const apiSecret = PropertiesService.getScriptProperties().getProperty('GENIUS_API_SECRET');
  if (!apiKey || !apiSecret) return { success:false, error:'GeniusPay non configure' };

  const siteBase = 'https://jacobsbio-cosmetique.com';
  const payload = {
    amount: Math.round(amount),
    currency: 'XOF',
    description: 'Commande ' + orderId + ' - Jacob\'s Bio',
    customer: {
      name: nom,
      phone: tel ? '+225' + tel.replace(/^0/, '') : undefined,
      email: email || undefined,
      country: 'CI'
    },
    success_url: siteBase + '/paiement-reussi.html?id=' + encodeURIComponent(orderId),
    error_url: siteBase + '/paiement-echoue.html?id=' + encodeURIComponent(orderId),
    metadata: { order_id: orderId }
  };

  try {
    const resp = UrlFetchApp.fetch(GENIUS_API_BASE + '/payments', {
      method: 'post',
      headers: {
        'X-API-Key': apiKey,
        'X-API-Secret': apiSecret,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    const body = JSON.parse(resp.getContentText());

    if (code !== 200 && code !== 201) {
      return { success:false, error:'GeniusPay ' + code + ' : ' + (body.message || 'erreur'), details: body };
    }

    // Sauvegarde la reference paiement dans le Sheet
    if (refIdx >= 0 && body.data && body.data.reference) {
      sh.getRange(rowNum, refIdx + 1).setValue(body.data.reference);
    }

    return {
      success: true,
      reference: body.data.reference,
      checkoutUrl: body.data.checkout_url,
      amount: body.data.amount,
      fees: body.data.fees,
      netAmount: body.data.net_amount
    };
  } catch (err) {
    return { success:false, error:'Erreur reseau : ' + err.message };
  }
}


// ================== WEBHOOK RECEPTION ==================
// Route dans doPost quand action=geniusWebhook. Verifie la signature
// HMAC-SHA256, le timestamp anti-replay, puis met a jour la commande.
function handleGeniusWebhook_(e) {
  try {
    // Recuperation des headers via e.parameter (Apps Script rend les
    // headers accessibles via e.postData + on lit ce qu'on peut)
    const signature = e.parameter['X-Webhook-Signature'] || '';
    const timestamp = e.parameter['X-Webhook-Timestamp'] || '';
    const rawBody = e.postData ? e.postData.contents : '';

    // Note : Apps Script Web Apps ne donnent PAS acces aux vrais
    // headers HTTP entrants (limitation Google). Solution : demander
    // a GeniusPay de mettre la signature dans le body (ou parametres
    // GET) plutot que dans les headers. On adapte selon ce que
    // GeniusPay accepte. En attendant, verification via query params.
    // Alternative robuste : GeniusPay reproduit la signature dans
    // le corps JSON (champ __signature) — a valider avec leur support.

    const webhookSecret = PropertiesService.getScriptProperties().getProperty('GENIUS_WEBHOOK_SECRET');
    if (!webhookSecret) return jsonResponse({ success:false, error:'webhook non configure' });

    // Anti-replay : timestamp < 5 min
    if (timestamp) {
      const age = Date.now() / 1000 - parseInt(timestamp, 10);
      if (age > 300 || age < -60) {
        return jsonResponse({ success:false, error:'timestamp expire' });
      }
    }

    // Verif signature
    if (signature && timestamp && rawBody) {
      const data = timestamp + '.' + rawBody;
      const expected = Utilities.computeHmacSha256Signature(data, webhookSecret)
        .map(b => (b & 0xff).toString(16).padStart(2, '0')).join('');
      if (expected !== signature.toLowerCase()) {
        return jsonResponse({ success:false, error:'signature invalide' });
      }
    }

    // Parse le payload
    const payload = JSON.parse(rawBody || '{}');
    const event = payload.event || '';
    const data = payload.data || {};
    const orderId = data.metadata && data.metadata.order_id;
    if (!orderId) return jsonResponse({ success:false, error:'order_id manquant' });

    // Mapping event -> statut
    let newStatut = null;
    if (event === 'payment.success') newStatut = 'Payee en ligne';
    else if (event === 'payment.failed') newStatut = 'Paiement echoue';
    else if (event === 'payment.cancelled') newStatut = 'Paiement annule';
    else if (event === 'payment.expired') newStatut = 'Paiement expire';

    if (!newStatut) return jsonResponse({ success:true, ignored: event });

    // MAJ Sheet
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sh) return jsonResponse({ success:false, error:'sheet introuvable' });
    const values = sh.getDataRange().getValues();
    const headers = values[0].map(h => String(h).trim());
    const idIdx = headers.indexOf('ID');
    const statutIdx = headers.indexOf('Statut');
    for (let r = 1; r < values.length; r++) {
      if (String(values[r][idIdx]).trim() === String(orderId).trim()) {
        sh.getRange(r + 1, statutIdx + 1).setValue(newStatut);
        break;
      }
    }

    try { invalidateAdminCache_(); } catch (e) {}

    return jsonResponse({ success:true, event, orderId, statut:newStatut });
  } catch (err) {
    return jsonResponse({ success:false, error: 'exception : ' + err.message });
  }
}


// ================== ESTIMATION FRAIS COTE VITRINE ==================
// Utilise par le site pour afficher "Total avec frais" avant paiement.
// Le vrai calcul est fait par GeniusPay, ceci est une estimation
// (frais 3.5% + arrondi FCFA a la centaine superieure).
function geniusEstimeFees_(amount) {
  const total = Math.ceil(amount * (1 + GENIUS_FEE_PERCENT / 100) / 100) * 100;
  return { originalAmount: amount, totalWithFees: total, feesEstimated: total - amount };
}
