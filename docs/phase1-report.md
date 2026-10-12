# Phase 1 : cœur et tests

Point de contrôle : `uv run pytest` (120 tests, FakeGateway) et `uv run pytest -m outlook`
(13 tests sur l'Outlook réel) sont verts.

## Ce qui est en place
- `core.Bridge` : point d'entrée unique (CLI, MCP, export) ; `Bridge.from_environment()` branche
  `ComGateway` et le `config.toml` de l'utilisateur.
- `com_gateway.ComGateway` : un seul thread STA (`outlook-com-sta`) et une file de requêtes ; seuls
  des modèles pydantic en sortent. Erreurs explicites : Outlook non lancé, nouvel Outlook,
  élévation, délai dépassé (boîte de dialogue ouverte).
- `mail`, `events`, `redact`, `timeutil`, `config`, `models`, `gateway`, `errors`.
- Garde lecture seule : appels interdits, écritures de propriétés COM (seule exception :
  `Items.IncludeRecurrences`), imports réseau, imports COM hors de `com_gateway.py`.

## Constats en cours de phase
1. **EntryID courts dans les `Table`** : en mode mis en cache, `GetTable` renvoie des EntryID
   *à court terme* (préfixe `EF`), valables pour la session seulement. On lit la colonne
   `PR_LONGTERM_ENTRYID_FROM_TABLE` (`0x66700102`), identique à `item.EntryID`.
2. **Arguments de date vers COM** : pywin32 attend l'heure murale locale étiquetée UTC (symétrique
   de la lecture). Un `datetime` naïf est décalé et `GetOccurrence` échoue. D'où
   `timeutil.to_com()`, inverse de `to_aware()`.
3. **`Parent` d'une occurrence** = l'élément maître (classe 26), pas le dossier.
4. **Fil de discussion** : pas de filtre DASL possible sur `PR_CONVERSATION_ID` (binaire). On
   retrouve le sujet de conversation (`thread-topic`) sur une table réduite, puis on filtre dessus
   (7,3 s → < 0,3 s sur la boîte de test). Le fil couvre Boîte de réception et Éléments envoyés.
5. Colonnes `Table` disponibles : aperçu `PR_PREVIEW` (255 car.) et `PR_SENDER_SMTP_ADDRESS`,
   d'où snippet et adresse SMTP sans ouvrir chaque élément.

## Écarts par rapport à la spec (à valider)
- **`calendar.py` → `events.py`** : un module `calendar` dans le paquet masque le module standard
  `calendar` (utilisé par `email`, `http.cookiejar`…) dès qu'un script du paquet est lancé
  directement, ce que prévoit `mcp dev src/outlook_bridge/mcp_server.py`.
- **Motif IBAN par défaut** : `\bFR\d{2}[ \d]{20,30}\b` avalait l'espace suivant et ignorait les
  lettres du numéro de compte. Remplacé par `\bFR\d{2}(?: ?[0-9A-Z]){23}\b`.
- **Ajouts au modèle** : `account` sur `MailSummary`/`Event`, `store_id` et `calendar` sur `Event`
  (exclusions par ID et multi-comptes), `body_truncated` sur `Event`.
- `get_event` accepte `start` pour viser une occurrence (toutes partagent l'EntryID de la série).
- `find_free_slots` compte comme occupés les événements masqués par mot-clé (seul le créneau sert,
  jamais le contenu) ; les comptes et calendriers exclus sont ignorés.
- Les rédactions sont appliquées avant troncature (marge de 200 caractères) pour ne jamais
  laisser un fragment de motif en fin de texte.

## Non couvert
- Journalisation (§6.4) : prévue avec la CLI et le serveur MCP (Phases 2 et 3).
- Réunion récurrente Exchange avec exceptions (aucune sur la boîte de test).
- `accounts_allowed` compare le `display_name` du store ; à revoir si un compte affiche un nom
  différent de son adresse.
