# Phase 2 : CLI, snapshots et journal d'audit

Point de contrôle : `uv run pytest` (660 tests, FakeGateway) et `uv run pytest -m outlook`
(18 tests sur l'Outlook réel, dont 5 pour la CLI) sont verts ; `ruff` propre.

## Ce qui est en place
- **CLI** (`outlook-bridge`, typer) : `status`, `mails`, `mail`, `events`, `slots`, `export`.
  Sortie texte en français, sans rich, nettoyée des séquences de contrôle ; `--json` en ASCII pur.
- **Snapshots** (`export`) : `inbox_digest.md`, `calendar_next14d.md` et `.json` dans
  `snapshot_dir` (ou `--out`). Toutes les requêtes Outlook passent avant toute écriture, puis
  écriture atomique en deux temps (fichiers temporaires `~$*.tmp` puis `os.replace`, réessayé si
  Drive ou OneDrive verrouille le fichier). En-tête `generated_at` et mention de contenu non fiable.
- **Journal d'audit** (`%LOCALAPPDATA%\outlook-bridge\logs\audit-AAAA-MM.jsonl`) : une ligne par
  appel d'outil (outil, paramètres nettoyés, nombre de résultats, durée, issue), écrite par le
  décorateur `@audited` posé sur les 9 méthodes du cœur qui correspondent aux outils MCP (§5.2).
- **Grammaire de dates** (`timeutil.parse_when`) et **formatage** (`fmt.py`) partagés par la CLI
  et l'export.

## Décisions validées au point de contrôle
1. **`mail <n°>`** au lieu du `mail <id>` littéral de la spec : le numéro renvoie à la dernière
   liste `mails` (`%LOCALAPPDATA%\outlook-bridge\state\last_mails.json`, IDs seulement). La forme
   complète `mail <EntryID> --store <compte>` reste disponible.
2. **`--to` au jour est inclusif** (`--to 16/10` inclut le 16) : par défaut `events` couvre 8 jours
   (aujourd'hui → même jour la semaine suivante) et `slots` 7 jours (demain → J+7).
3. **Contenu des snapshots** dans le dossier synchronisé : sujets, expéditeurs, extraits de 200
   caractères au plus, lieux, organisateur, noms des participants (JSON seulement), créneaux libres.
   Jamais de corps, d'EntryID/StoreID ni de lien de visio (le JSON n'indique que `has_online_meeting`).
4. **Journal d'audit en mode « fail-open »** : s'il ne peut pas être écrit, un avertissement sur
   stderr, et les appels continuent.

## Garanties de confidentialité (par construction)
- Journal : les IDs, comptes et dossiers non standard sont remplacés par une empreinte
  (`#` + 10 caractères hexadécimaux) ; le texte libre (`query`, `sender`) par sa seule longueur
  (`<texte:N>`) ; une erreur par son nom de classe. Le journal ne reçoit jamais un résultat, seulement
  un nombre. Des tests « canari » le vérifient pour chaque outil.
- Les messages d'exception inattendus ne sont jamais affichés (ils peuvent contenir du contenu de
  mail) : `Erreur interne (<Classe>)`, trace complète seulement avec `--debug`.
- Aucun chemin utilisateur absolu en sortie : `%APPDATA%`, `%LOCALAPPDATA%` ou `%USERPROFILE%`.
- `export` refuse d'écrire dans un dépôt git ; `.gitignore` couvre snapshots, journal, état CLI.
- Le texte non fiable est échappé dans le Markdown des snapshots (ni lien, ni image, ni HTML, ni
  front matter injectable).

## Codes de sortie
0 succès · 1 erreur interne · 2 usage (date, plage, numéro, compte) · 3 Outlook indisponible ·
4 introuvable ou exclu · 5 configuration · 6 écriture des snapshots · 130 interruption.

## Écarts et précisions
- Les messages de validation de la Phase 1 disent « doit valoir au moins 1 » : le caractère `≥`
  n'existe ni en cp1252 ni en cp850 (consoles Windows).
- `14h` seul est un décalage (« dans 14 heures ») ; une heure de la journée s'écrit `14:00` ou
  `14h00`. Après une date, `16/10 14h` est bien une heure.
- PowerShell 5.1 décode la sortie redirigée (`> fichier`, `| Out-File`) avec la page OEM : les
  accents peuvent y être mal rendus (sans plantage). `--json` n'est pas concerné (ASCII). Solution :
  `[Console]::OutputEncoding = [Text.Encoding]::UTF8`.
- `snapshot_dir` doit être un chemin absolu écrit en chaîne littérale TOML (guillemets simples),
  sinon `\M` dans `G:\Mon Drive` est une séquence d'échappement invalide.

## Revue adversariale
Revue en 4 axes (exactitude, confidentialité, robustesse Windows, qualité des tests), chaque
constat confirmé par deux vérificateurs indépendants : 20 confirmés, 8 partagés, tous corrigés
avec un test de non-régression (819 tests). Les plus importants :
- **Liens de visio** : ils passaient encore dans les snapshots par le champ `location` ; toute URL
  y est désormais remplacée par `[lien retiré]` (JSON et Markdown).
- **Extraits** : la passerelle coupait l'aperçu à 200 caractères avant la rédaction, si bien qu'un
  IBAN à cheval sur la coupe échappait au masquage. Elle garde maintenant une marge, et la fin d'un
  aperçu déjà tronqué par Outlook (environ 255 caractères) est retirée.
- **Configuration invalide** : le message ne cite plus les valeurs (expéditeurs, mots-clés, motifs
  exclus), seulement le champ ; fichiers ANSI, UTF-16 ou avec BOM gérés (code de sortie 5).
- **`mail`** : les délimiteurs du contenu portent un jeton aléatoire, qu'un corps de mail ne peut
  pas imiter pour faire passer du texte pour la sortie de l'outil.
- **Dossier ou calendrier exclu** : même réponse qu'un dossier ou calendrier inconnu (on ne peut
  plus sonder `folders_excluded`).
- **Événements sur plusieurs jours** : listés sous chaque jour couvert, marqués `(suite)`, pour
  qu'un congé ne fasse jamais apparaître des jours « libres ».
- **Tests sur le vrai Outlook** : n'assertent que des booléens ou des nombres calculés à part
  (vérifié statiquement), et suppriment les snapshots réels qu'ils écrivent.
- Écritures concurrentes : fichiers temporaires à nom unique (export, état CLI) et verrou sur le
  journal JSONL entre processus.

## Non couvert
- Planification de l'export (`schtasks`, toutes les 30 min) : Phase 6.
- Rétention du journal : fichiers mensuels, purge manuelle (à revoir en Phase 6).
