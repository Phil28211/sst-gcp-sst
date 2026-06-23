# Accueil SST - GCP Construction

Mini-application web statique pour GitHub Pages.

## Inclus
- Formulaire mobile
- Photos de cartes
- Signature tactile
- Génération PDF
- Téléchargement du PDF
- Option courriel automatique avec EmailJS

## Installation GitHub Pages
1. Créer un compte GitHub.
2. Créer un nouveau repository : `gcp-accueil-sst`.
3. Ajouter `index.html`, `style.css`, `script.js`, `README.md`.
4. Aller dans `Settings` > `Pages`.
5. Source : `Deploy from a branch`.
6. Branch : `main`, Folder : `/root`.
7. Save.
8. Le site sera disponible sur un lien du type :
   `https://votre-utilisateur.github.io/gcp-accueil-sst/`

## Courriel automatique
GitHub Pages ne peut pas envoyer un courriel seul. Pour envoyer automatiquement le PDF, configurer EmailJS dans `script.js` :
- EMAIL_ENABLED = true
- EMAILJS_PUBLIC_KEY
- EMAILJS_SERVICE_ID
- EMAILJS_TEMPLATE_ID
- EMAIL_TO

Dans le template EmailJS, prévoir une pièce jointe avec le paramètre `pdf_attachment`.

## Sous-domaine
Pour `sst.gcpconstruction.ca`, configurer un Custom Domain dans GitHub Pages et un enregistrement DNS CNAME.
