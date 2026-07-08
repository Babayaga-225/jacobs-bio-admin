# Jacob's Bio · Admin CRM

Interface d'administration pour la gestion des commandes Jacob's Bio Cosmétique.

## Fonctionnalités (Phase 2A)

- **Login** protégé par mot de passe (session 24h)
- **Dashboard** temps réel : CA jour/mois, commandes en attente, stock bas
- **Gestion commandes** : filtre (aujourd'hui/en attente/7 jours/toutes), actions rapides
  - Confirmer / Expédier / Livrée / Annuler
  - Contact WhatsApp direct
- **Stock bas** : liste des produits sous le seuil de 5 unités

## Stack

- HTML/CSS/JS pur, single-file SPA
- Backend : Google Apps Script (Web App)
- Auth : mot de passe partagé + token HMAC-SHA256 (24h TTL)

## Déploiement Vercel

1. Fork/clone ce repo
2. Nouveau projet Vercel > importer le repo GitHub
3. Framework preset : *Other* (static)
4. Deploy

## Configuration

L'URL de l'Apps Script Web App est en dur dans `index.html` (constante `API_URL`).
Le mot de passe est configuré côté Apps Script dans les **Propriétés du script** :

- `ADMIN_PASSWORD` : le mot de passe partagé
- `ADMIN_SECRET` : clé secrète pour signer les tokens

Voir `APPS_SCRIPT_A_AJOUTER.js` pour le code backend à coller dans le projet Apps Script.

## Local

```powershell
python -m http.server 8766
# Puis ouvrir http://localhost:8766/
```

## Roadmap

- **Phase 2B** — Contrôle promos & bandeau (toggle promo, texte popup, marquee)
- **Phase 2C** — CRUD produits (ajout/édition/photos)

---

MyComm Platform · Bouaké · 2026
