# Kit Séjour — l'appli des vacances en tribu

Une petite appli web pour organiser une semaine de vacances à plusieurs familles
dans une grande maison : planning des repas, dépenses partagées, comptes de fin
de séjour, répartition des chambres, wifi, liste de courses, participants.

Pensée pour être utilisée sur smartphone via un simple lien WhatsApp, par des
utilisateurs de 7 à 77 ans. Pas de compte à créer, pas d'appli à installer :
quiconque a le lien voit et modifie tout (c'est assumé, c'est de la famille).

Chaque séjour = une copie indépendante de ce kit, gratuite, hébergée sur votre
compte Google (Apps Script + un Google Sheet pour les données).

## Ce que fait l'appli

- **Réglages** : titre du séjour, lieu, dates d'arrivée et de départ,
  logements et chambres avec leurs couchages. Tout est éditable directement
  dans l'appli, rien à coder.
- **Repas** : un planning déjeuner/dîner généré automatiquement d'après les
  dates (dîner seul le premier jour, déjeuner seul le dernier). Chaque service
  est pris par un ménage ou une équipe libre, avec menu facultatif. Absents
  signalables par repas (défalqués des comptes) et invités de passage
  ajoutables par repas (leurs parts à la charge du ménage qui invite).
- **Dépenses** : saisie rapide, facture en photo ou PDF (facultatif), liste de
  courses partagée, partage WhatsApp. Quatre clefs de répartition : aux parts
  (tout le monde), entre certains ménages, entre les amateurs d'apéro/alcool,
  ou par personne à prix par tête (ex : bateau à 25 € par personne, on coche
  qui participe).
- **Comptes** : répartition au nombre de parts (1 part par adulte, ½ part par
  enfant de moins de 10 ans, 0 pour les bébés et invités), remboursements de
  fin de séjour suggérés automatiquement. Option : prorata des nuits de
  présence, pour les groupes où chacun arrive et repart quand il veut.
- **Sauvegarde et export** (onglet Réglages) : export de tout le séjour en
  classeur Google Sheets (téléchargeable en Excel ou CSV), récap complet par
  mail avec la sauvegarde des données en pièce jointe, copie/restauration de
  sauvegarde, et « repartir à zéro » pour enchaîner un nouveau séjour sur la
  même appli (l'ancien état reste archivé dans le classeur de données).
- **Famille** : participants, ménages (les unités qui paient : ajout,
  renommage, suppression), parts, et dates d'arrivée/départ de chacun.
- **Chambres** : répartition par logement avec alerte de surcapacité, encart
  wifi par bâtiment (mot de passe partagé + QR de connexion).
- **Admin** (icône ⋯) : alertes mail à chaque modification, par catégorie,
  protégées par mot de passe.

## Dupliquer pour votre séjour

Prérequis : un compte Google, [Node.js](https://nodejs.org) installé, et
`clasp`, l'outil en ligne de commande officiel de Google Apps Script (installé
automatiquement par `npx` à la première commande).

### 1. Récupérer le kit

```
git clone https://github.com/joachimpomme-ctrl/kit-sejour-famille.git mon-sejour
cd mon-sejour
```

### 2. Créer le projet Apps Script

```
npx @google/clasp login          # une seule fois par machine
npx @google/clasp create --type webapp --title "Mon séjour"
npx @google/clasp push -f
```

`create` écrit un fichier `.clasp.json` local (l'identifiant de VOTRE copie,
il n'est pas versionné). Si `create` propose d'écraser `appsscript.json`,
refusez, ou re-poussez ensuite avec `push -f`.

### 3. Autoriser le script (une seule fois, dans l'éditeur)

```
npx @google/clasp open-script
```

Dans l'éditeur qui s'ouvre : sélectionnez la fonction **initSejour** et
cliquez **Exécuter**. Google demande les autorisations (Sheets, Drive, Gmail) :
acceptez. Cette étape crée le classeur de données et le dossier des factures
dans votre Drive.

Important : si vous sautez cette étape, l'appli affichera une page
d'autorisation à tous vos invités au lieu de fonctionner.

### 4. Déployer en appli web

Toujours dans l'éditeur : **Déployer → Nouveau déploiement → Application
Web**, avec :
- Exécuter en tant que : **Moi**
- Qui a accès : **Tout le monde** (bien choisir la variante anonyme, pas
  « avec compte Google »)

Copiez l'URL en `.../exec` : c'est le lien à partager. Testez-la en navigation
privée : l'appli doit s'afficher sans demander de connexion.

Après toute modification du code : `npx @google/clasp push -f` puis, dans
l'éditeur, **Déployer → Gérer les déploiements → ✏️ → Nouvelle version**.
Ne créez jamais un deuxième déploiement : l'URL changerait.

### 5. (Facultatif) Une jolie URL sans bandeau Google, via Vercel

L'URL `/exec` affiche un bandeau « créée par un utilisateur Apps Script ».
Pour une URL propre type `mon-sejour.vercel.app` sans bandeau :

1. Dans `wrapper/index.html`, remplacez les deux occurrences de
   `URL_EXEC_A_COLLER` par votre URL `/exec`.
2. Personnalisez le nom affiché (balises `<title>`, `apple-mobile-web-app-title`
   et `wrapper/manifest.webmanifest`).
3. (Facultatif) Régénérez les icônes avec votre initiale :
   `powershell -ExecutionPolicy Bypass -File outils/icones.ps1 -Lettre "M"`
   (sous Windows ; sinon gardez les icônes fournies).
4. Déployez : `npx vercel deploy --prod --yes --cwd wrapper`
   (compte [Vercel](https://vercel.com) gratuit requis, `npx vercel login` la
   première fois).

L'appli s'installe alors comme une vraie appli : iPhone → Safari → Partager →
« Sur l'écran d'accueil » ; Android → Chrome → menu ⋮ → « Ajouter à l'écran
d'accueil ».

### 6. Paramétrer le séjour

Ouvrez le lien : l'appli démarre sur l'onglet **Réglages**. Donnez un titre,
les dates et les logements, puis créez les ménages et ajoutez les participants
dans **Famille**. C'est prêt : partagez le lien dans le groupe WhatsApp.

Admin : icône ⋯ en haut. Identifiant = votre adresse Gmail, mot de passe par
défaut `vacances` — **changez-le tout de suite** dans le panneau ⚙.

## Architecture (pour les curieux)

- `Code.gs` — serveur Apps Script. Les données vivent dans un JSON unique en
  cellule A1 d'un Google Sheet créé automatiquement. Toutes les écritures
  passent par un verrou (`LockService`) : opérations atomiques, pas de conflit
  entre téléphones.
- `Index.html` — client complet en vanilla JS, sans framework ni build.
  Re-rendu complet à chaque mise à jour, ~1 s de latence par écriture (c'est
  Apps Script, c'est normal, et c'est gratuit).
- `wrapper/` — page statique Vercel qui affiche l'appli en iframe plein écran
  (supprime le bandeau Apps Script) + manifeste PWA.
- Pas d'authentification des utilisateurs, pas d'historique : le Sheet fait
  office de sauvegarde consultable. Simplicité avant tout.

## Limites assumées

- Quiconque a le lien peut tout modifier (public familial).
- Pas d'édition du libellé ou du montant d'une dépense : supprimer puis
  ressaisir. L'affectation (catégorie, repas, ménages ciblés), elle, se
  modifie en tapant dessus dans la liste.
- Supprimer une dépense ne supprime pas sa facture du Drive (archive).
- L'appli nécessite une connexion : pas de mode hors ligne.

## Licence

MIT. Faites-en bon usage, et bonnes vacances.
