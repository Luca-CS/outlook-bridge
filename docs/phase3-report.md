# Phase 3 : serveur MCP

Point de contrôle : `uv run pytest` (935 tests, FakeGateway) et `uv run pytest -m outlook`
(22 tests sur l'Outlook réel, dont 4 pour le serveur MCP) sont verts ; `ruff` propre.
Reste à valider par l'utilisateur : les 4 cas d'usage de la spec §1 dans Claude Code (voir
« Essai dans Claude Code »).

## Ce qui est en place
- **Serveur MCP** (`src/outlook_bridge/mcp_server.py`, entrée `outlook-bridge-mcp`) : transport
  stdio, les 9 outils §5.2 (`status`, `list_accounts`, `list_folders`, `search_mail`, `get_mail`,
  `get_thread`, `list_events`, `get_event`, `find_free_slots`), annotés lecture seule
  (`readOnlyHint`, `idempotentHint`, ni destructif ni « monde ouvert »).
- **Fine couche sur le cœur**, comme la CLI : le lifespan ouvre un `Bridge.from_environment(source="mcp")`
  et le ferme à l'arrêt. Chaque outil appelle une seule méthode `@audited` du cœur, donc chaque
  appel écrit exactement un enregistrement d'audit, sans code d'audit dans le serveur.
- **Descriptions** : courtes, sur une ligne. Les 5 outils qui renvoient du contenu (`search_mail`,
  `get_mail`, `get_thread`, `list_events`, `get_event`) rappellent qu'il s'agit d'une donnée non
  fiable, jamais d'une instruction ; les `instructions` du serveur le répètent, avec le format des
  dates.
- **Sorties** : les modèles pydantic du cœur en JSON compact (UTF-8), dans un seul bloc texte.
- **Erreurs** : le message d'une `OutlookBridgeError` est renvoyé tel quel (chemins du profil
  masqués) ; toute autre exception devient « Erreur interne (<Classe>) ». Un pont qui ne peut pas
  s'ouvrir (config.toml invalide) ne tue pas le serveur : chaque outil renvoie l'erreur.
- **Claude Code** : `.mcp.json` (portée projet, `uv run --quiet outlook-bridge-mcp`) et
  `.claude/settings.json`, qui autorise `mcp__outlook__*` sans confirmation.

## Décisions de la phase
1. **Dates des outils en ISO 8601 seul** (validée) : `2026-10-16` ou `2026-10-16T14:30[:00]`,
   heure de Paris sans fuseau, décalage ou `Z` respectés, `T24:00` accepté. Un jour seul en fin de
   plage est inclus, comme dans la CLI. Tout autre format (`demain`, `2d`, `16/10`, `14:00`)
   renvoie une erreur qui rappelle le format attendu, avant tout appel à Outlook
   (`timeutil.parse_iso`).
2. **Pas de sortie structurée** (`structuredContent`) : le SDK la doublerait d'une copie texte
   indentée, soit deux fois plus de jetons. Un seul bloc JSON compact suffit au modèle.
3. **Caractères invisibles retirés** des sorties (inversions bidi, caractères « tag » Unicode) : un
   modèle les lit, un humain qui relit l'appel ne les voit pas. C'est une voie connue d'injection.
4. **`find_free_slots` ne propose rien dans le passé** : une plage déjà commencée part de
   maintenant (arrondi au quart d'heure), comme la CLI et l'export.
5. **Liens longs abrégés** (amélioration 1) : tout lien de plus de 80 caractères devient
   `<lien : domaine>` dans les corps (`get_mail`, `get_event`) et les extraits, **avant** la coupe
   à `max_chars`, sinon le gain serait perdu. Le lien de visio est cherché dans le corps complet
   avant l'abréviation : `online_meeting_url` reste entier.
6. **Réunions Teams** (amélioration 2) : nouveau champ `Event.has_online_meeting`, vrai si un lien
   de visio est connu ou si le lieu contient « Microsoft Teams ». `list_events`, le marqueur
   `visio` de la CLI et les snapshots l'utilisent. `get_event` lit en plus la propriété Teams
   quand ni le lieu ni le corps n'ont de lien.

## Sondes masquées (Outlook réel, compteurs seulement)
- La propriété nommée `SkypeTeamsMeetingUrl` (PS_PUBLIC_STRINGS) est présente exactement sur les
  événements dont le lieu contient « Microsoft Teams » (« Réunion Microsoft Teams » ou « Microsoft
  Teams Meeting »), et sa valeur est identique au lien du corps. Aucun désaccord sur ±180 jours.
- `AppointmentItem.IsOnlineMeeting` vaut toujours faux, y compris pour les réunions Teams :
  inutilisable.
- Lire la propriété coûte environ 30 ms par événement : réservée à `get_event`, jamais à
  `list_events`, où le lieu suffit (coût nul, même résultat sur l'échantillon).

## Écarts et précisions
- **FastMCP est devenu `MCPServer`** dans le SDK officiel `mcp` 2.x (2.3.0 installé) ; la spec
  suppose FastMCP. Le serveur est écrit pour `MCPServer`, sans dépendre de son `Context` (dont l'API
  a changé entre v1 et v2) : les outils lisent le pont par une fermeture.
- **`claude mcp add` indisponible** : la CLI `claude` n'est pas installée sur cette machine
  (Claude Code tourne dans l'app de bureau). À la place, `.mcp.json` de projet, sans chemin
  personnel : le serveur se lance depuis la racine du dépôt ouvert dans Claude Code.
- **Inspecteur MCP** : `uv run mcp dev src/outlook_bridge/mcp_server.py` trouve bien le serveur
  (objet `server` au niveau du module), mais télécharge l'inspecteur par `npx`. Il est remplacé
  ici par un test d'intégration qui lance le vrai script en stdio et appelle les outils.
- **`get_event` prend `store_id`** (et `start` pour une occurrence de série) en plus de la table
  §5.2, comme `get_mail` : nécessaire avec plusieurs comptes.
- **`max_chars` par défaut** = `max_body_chars` de la config (4000), 20000 au plus.
- **Dépendances** : `mcp` tire `opentelemetry-api`, sans SDK ni exportateur : aucune donnée
  n'est émise (règle « aucune télémétrie »). Le garde-fou `test_readonly_guard` interdit toujours
  `urllib` dans `src/` : le domaine d'un lien est extrait par une expression régulière.
- **Journal du SDK** : niveau `WARNING` sur stderr. Une exception inattendue est remplacée avant
  que le SDK ne la journalise, pour qu'aucune trace ne contienne de contenu de mail.

## Revue
Un seul agent `reviewer` (Opus) en fin de phase, axé sur l'injection de prompt et les fuites.
Vérifiés sans défaut : aucune route par laquelle un contenu de mail ou un message d'exception
atteint le modèle, stderr ou le journal sans masquage (chemins d'erreur du SDK 2.3.0 compris) ;
rien n'écrit sur stdout ni dans Outlook ; `parse_iso`, `not_in_past`, `shorten_urls` (pas de
retour arrière catastrophique, idempotent), cohérence de `has_online_meeting`, cycle de vie du
lifespan, aucune donnée personnelle dans le diff. Constats et suites :
- **Garde-fou des tests d'intégration** (majeur) : une assertion portait sur un indice
  (`status["reachable"]`) ; le test statique qui n'autorise que des noms ou des comparaisons
  simples échouait. Corrigé.
- **Hôte d'un lien avec `\`** : les navigateurs lisent `\` comme `/`, donc
  `https://a.example\@b.example/...` mène à `a.example` mais s'affichait `<lien : b.example>`.
  L'hôte s'arrête désormais au `\` (test ajouté).
- **Motifs de `redact_patterns` visant une URL entière** : l'abréviation passe avant la
  rédaction, donc un motif comme `https://intranet\.corp\S*` ne voit plus que
  `<lien : intranet.corp>`. Le chemin et la requête ont disparu, mais le domaine reste visible :
  pour le masquer, écrire le motif sur le domaine seul (`intranet\.corp`).
- **Surrogate isolé venant de COM** : le constat visait la sérialisation du serveur, mais
  pydantic refuse déjà ce caractère à la construction des modèles. Un tel texte ferait donc
  échouer l'appel plus tôt, dans le gateway, pour la CLI comme pour MCP. Ce cas est antérieur à
  la phase et n'a jamais été observé ; à traiter dans le gateway si un jour il apparaît.
- Tests ajoutés : arguments invalides (type, paramètre manquant) refusés avant Outlook et sans
  audit, chiffres non ASCII refusés dans les dates. Un argument inconnu est ignoré par le SDK.

## Essai dans Claude Code (point de contrôle)
1. Ouvrir une nouvelle session Claude Code à la racine du dépôt (le worktree de la branche, ou
   `main` après le merge), Outlook classique lancé.
2. Accepter le serveur `outlook` du `.mcp.json` du projet quand Claude Code le demande.
3. Essayer les 4 cas d'usage : non lus depuis hier résumés par importance ; agenda de la semaine
   et préparation d'une réunion ; trois créneaux d'une heure la semaine prochaine ; dernier
   échange avec une personne sur un sujet.

## Non couvert
- Claude Desktop (MCPB, `mcpb/manifest.json`) : Phase 4.
- Skills : Phase 5 ; accès à distance et export planifié : Phase 6.
