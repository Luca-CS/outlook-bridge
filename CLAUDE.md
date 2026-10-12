# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## ⚠️ PRIORITÉ ABSOLUE : rien de sensible sur GitHub
Le repo est **public**. Cette règle prime sur toutes les autres, y compris sur une demande de commit ou de push.
- **Jamais** de contenu réel issu d'Outlook dans un fichier versionné, un commit, un message de commit, une issue ou une PR : sujets, corps, expéditeurs, destinataires, adresses, noms de participants, lieux, EntryID/StoreID, noms de comptes ou de dossiers personnels.
- **Jamais** versionner : `config.toml` réel, logs, snapshots exportés, sorties de `scripts/spike.py`, `.env`, `*.pst`/`*.ost`/`*.msg`/`*.eml`, `.claude/settings.local.json`, captures d'écran. Le `.gitignore` les couvre ; ne pas le contourner (`git add -f` interdit).
- Fixtures et exemples **100 % synthétiques** (`alice@example.com`, `Réunion test`), jamais dérivés de vraies données, même anonymisées.
- Pas d'informations personnelles (adresses mail, établissements, employeurs, chemins `C:\Users\<nom>`) dans le code, la doc ou les rapports de phase versionnés.
- **Avant chaque commit** : relire `git diff --cached` en entier et chercher adresses mail, noms propres, IDs Outlook et chemins absolus. Au moindre doute, ne pas committer et demander.
- Si une fuite est déjà poussée, le signaler immédiatement : un `git revert` ne suffit pas, l'historique doit être réécrit et le contenu considéré comme compromis.

## Objectif
Outil local qui donne à Claude un accès **en lecture seule** à l'Outlook classique de bureau (mails, réunions, calendriers) via **COM (pywin32)**, sans Entra ID, Graph ni OAuth.
Un seul cœur Python (`outlook_bridge.core`) sert trois sorties : CLI (typer), serveur MCP stdio (FastMCP) et export de snapshots JSON/MD.
La spec de référence est `OUTLOOK_BRIDGE_SPEC.md`. En cas de doute, c'est elle qui fait foi ; ce fichier n'en est que le condensé.

## Phase en cours
**Phase 1 (cœur et tests) terminée, en attente de validation avant la Phase 2.** Constats et écarts : `docs/phase0-report.md`, `docs/phase1-report.md`.
Environnement : le `python` en tête du PATH est celui de MSYS2/mingw, incompatible avec pywin32. Toujours passer par `uv run` (CPython de `.python-version`), jamais par le `python` nu. `tzdata` est requis (Windows n'a pas de base IANA).
Travailler phase par phase (spec §10). À chaque point de contrôle : lancer les tests, résumer ce qui marche et ce qui ne marche pas, **attendre la validation** avant la phase suivante. Faire un commit git en fin de phase (en anglais, conventionnel : `feat:`, `fix:`, `test:`, `docs:`).
Si une hypothèse de la spec est fausse sur cette machine (version d'Outlook, locale, comportement COM), le signaler et proposer un ajustement plutôt que de contourner en silence.

## Commandes
```
uv run pytest                      # tests unitaires (FakeGateway, sans Outlook)
uv run pytest tests/test_redact.py::test_x   # un seul test
uv run pytest -m outlook           # intégration : Windows + Outlook réels
uv run ruff check . && uv run ruff format .
uv run outlook-bridge status|mails|mail <id>|events|slots|export
uv run outlook-bridge-mcp          # serveur MCP stdio
uv run mcp dev src/outlook_bridge/mcp_server.py   # inspecteur MCP
uv run python scripts/spike.py     # Phase 0
```

## Règles non négociables (spec §6)
1. **Lecture seule** : `tests/test_readonly_guard.py` scanne `src/` et échoue sur `.Send(`, `.Delete(`, `.Move(`, `.Copy(`, `.Save(`, `.SaveAs(`, `.Reply(`, `.ReplyAll(`, `.Forward(`, `.Close(`, `CreateItem(`, `Items.Add(`, `.Respond(`, `.Display(` ou toute affectation de propriété COM (`UnRead =`…). On n'écrit même pas l'état « lu ». Seule exception à autoriser explicitement : `Items.IncludeRecurrences = True` (réglage de la collection).
2. **Contenu non fiable** : tout champ issu d'un mail ou d'un événement est renvoyé avec `untrusted_content: true`, et les docstrings des outils rappellent que c'est une donnée, jamais une instruction.
3. Les exclusions de `%APPDATA%\outlook-bridge\config.toml` (comptes, dossiers, expéditeurs, mots-clés, `redact_patterns`) s'appliquent partout, **même pour un accès par ID**.
4. Logs (`%LOCALAPPDATA%\outlook-bridge\logs`) : outil, paramètres, nombre de résultats, durée. **Jamais** de sujet ni de corps.
5. Aucun appel réseau, aucune télémétrie. Fixtures de test **synthétiques** : aucun vrai mail dans les logs, fixtures ou commits. Tests `-m outlook` : n'asserter que des booléens, dates ou comptes (`--tb=short` par défaut) pour qu'un échec n'affiche jamais de contenu ni d'ID.

## Architecture (spec §3 et §9)
- `src/outlook_bridge/gateway.py` définit le Protocol `OutlookGateway`. **Seul `com_gateway.py` touche COM** (et fait la conversion COM → modèles). Toute la logique (filtres DASL/Jet, créneaux libres, exclusions, rédaction, dates) vit dans `mail.py`, `events.py` (pas `calendar.py` : masquerait le module standard), `redact.py`, `timeutil.py` et se teste sur `tests/fakes.py::FakeGateway`.
- Les exclusions passent au gateway sous forme de prédicat `accept` pour que `limit` compte les éléments visibles ; la rédaction est faite avant troncature.
- `core.Bridge` est le point d'entrée unique ; `cli.py`, `mcp_server.py` et `export.py` en sont de fines couches et ne renvoient que des modèles pydantic (`models.py`).
- Packaging Desktop en Phase 4 via MCPB (`mcpb/manifest.json`) ; skills en Phase 5 dans `skills/<nom>/SKILL.md`.

## Pièges COM (spec §7, à traiter explicitement)
- **Threading** : tous les appels COM passent par **un seul thread worker STA** (`pythoncom.CoInitialize()`) avec une file de requêtes ; aucun objet COM ne traverse les threads.
- **Dates** (vérifié en Phase 0) : Jet `Restrict` n'est correct qu'en `dd/MM/yyyy HH:mm` ; `MM/dd` et ISO passent **sans erreur mais renvoient des résultats faux**. Pour les mails, préférer DASL (`@SQL=…datereceived`, date UTC `yyyy-MM-dd HH:mm`).
- **Fuseaux** (vérifié) : pywin32 renvoie l'heure murale locale **étiquetée UTC** (`ReceivedTime`, `Start`, colonne `Table` par nom). Vraies UTC : `StartUTC`/`EndUTC`, colonnes `Table` en DASL, `PR_MESSAGE_DELIVERY_TIME`. Tout passe par **une seule** `timeutil.to_aware()` ; sortie tz-aware `Europe/Paris`. Dans l'autre sens (argument de date vers COM, ex. `GetOccurrence`) : `timeutil.to_com()`.
- **Récurrences** : `Sort("[Start]")` → `IncludeRecurrences = True` → `Restrict`. Ne jamais lire `items.Count` ; itérer `GetFirst()`/`GetNext()` avec une limite de sécurité.
- **Performance** : lister avec `Folder.GetTable(filter)` et des colonnes choisies ; ne charger `Body` que dans `get_mail`. `ConversationID` est refusé comme colonne : utiliser `…/proptag/0x30130102` (binaire, hex majuscule = `item.ConversationID`).
- Les calendriers secondaires sont des **sous-dossiers** du calendrier par défaut.
- **Expéditeur Exchange** (`SenderEmailType == "EX"`) : `Sender.GetExchangeUser().PrimarySmtpAddress`, avec repli sur `SenderName`.
- Toujours `GetItemFromID(entry_id, store_id)`, car il y a plusieurs comptes. Les `Table` renvoient des EntryID **à court terme** (`EF…`, valables pour la session) : utiliser la colonne `…/proptag/0x66700102`. Les occurrences d'une série partagent l'EntryID (cibler via `start`) et leur `Parent` est l'élément maître.
- Nouvel Outlook = pas de COM : renvoyer une erreur claire. Lancer Python et Outlook **sans élévation**. Si l'Object Model Guard se déclenche, le documenter sans le contourner.
