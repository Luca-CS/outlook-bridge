# Phase 0 : rapport du spike de faisabilité

Script : `scripts/spike.py` (sortie masquée par défaut, `--show` pour voir les valeurs en local).
Verdict : **GO**. Toutes les hypothèses bloquantes sont levées.

## Environnement constaté
- Windows 10, locale `fr-FR` (dates `dd/MM/yyyy`), fuseau `Romance Standard Time`.
- Outlook classique 16.0 (Microsoft 365, Click-to-Run), nouvel Outlook non lancé, processus non élevés.
- Un compte Exchange, en **mode mis en cache** (`ExchangeConnectionMode = 700`, `Store.IsCachedExchange = True`).
- uv 0.13 installé via winget ; CPython 3.14 (64 bits) ; pywin32 312.

## Résultats
| Point | Résultat |
|---|---|
| Connexion `Outlook.Application` | OK (`Dispatch` se rattache à l'instance lancée) |
| Comptes / stores | OK |
| 5 derniers mails (sujet, expéditeur, date) | OK |
| Expéditeur Exchange (`EX`) → SMTP via `GetExchangeUser().PrimarySmtpAddress` | OK |
| Object Model Guard | **Aucun avertissement** (adresses, `Body`, `PropertyAccessor` : aucun appel lent) |
| `Restrict` Jet sur la locale FR | Seul **`dd/MM/yyyy HH:mm`** est correct. `MM/dd/yyyy` et ISO `yyyy-MM-dd` sont **acceptés sans erreur mais faux** (résultats silencieusement erronés) |
| `Restrict` DASL `urn:schemas:httpmail:datereceived` | OK avec une date UTC `yyyy-MM-dd HH:mm` |
| `GetTable` + colonnes choisies | OK ; `GetItemFromID(entry_id, store_id)` OK |
| Récurrences (`Sort` → `IncludeRecurrences` → `Restrict`) | OK : séries annuelles développées en occurrences |

## Écarts par rapport à la spec (ajustements proposés)
1. **`tzdata` obligatoire** : Windows n'a pas de base IANA, `ZoneInfo("Europe/Paris")` échoue sans le paquet `tzdata`. Ajouté aux dépendances (à reporter dans spec §4).
2. **Fuseaux (§7.6) confirmé** : pywin32 renvoie l'**heure murale locale étiquetée UTC** (`TimeZoneInfo('GMT Standard Time', True)`) pour `ReceivedTime`, `Start`, `End`, et pour la colonne `ReceivedTime` d'une `Table`. Les sources réellement UTC sont : `StartUTC`/`EndUTC`, `PR_MESSAGE_DELIVERY_TIME` via `PropertyAccessor`, et les colonnes de `Table` référencées par nom DASL (`urn:schemas:httpmail:datereceived`). Proposition pour `to_aware()` : partir d'une valeur **UTC** quand elle existe et convertir vers `Europe/Paris` (aucune ambiguïté au changement d'heure) ; sinon remplacer le tzinfo par `Europe/Paris` (ambigu une heure par an, le dernier dimanche d'octobre entre 2h et 3h).
3. **`ConversationID` refusé comme colonne de `Table`** (`Paramètre incorrect`). Utiliser le tag `PR_CONVERSATION_ID` (`http://schemas.microsoft.com/mapi/proptag/0x30130102`, binaire 16 octets) : son hex majuscule est égal à `item.ConversationID` (vérifié).
4. **Calendrier par défaut vide sur les 31 prochains jours**, et sans série récurrente : la vérification « réunion récurrente dans les 7 jours » a été faite sur un sous-calendrier (jours fériés, séries annuelles) et sur des événements passés du calendrier par défaut. Les calendriers secondaires sont des **sous-dossiers** du calendrier par défaut : `list_events(calendar=…)` devra les parcourir.
5. Mode mis en cache : à signaler dans `status` (§7.10), comme prévu.

## Non testé ici
- Réunion récurrente Exchange avec exceptions (occurrence déplacée ou annulée) : aucune sur ce compte.
- Comportement au changement d'heure du 25/10 pour des événements horaires (seulement des journées entières observées, décalage +1h correct).
- DASL avec une heure locale par erreur : non discriminant sur cet échantillon (pas de mail entre 22h et minuit UTC au seuil).
