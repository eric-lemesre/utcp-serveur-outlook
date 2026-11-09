# Outlook : couverture du manuel UTCP

Manuel `outlook` 0.1.0, UTCP 1.0.2, 49 outils, protocole `http` seul, Microsoft Graph v1.0. L'URL de base `${GRAPH_BASE}` (défaut `https://graph.microsoft.com/v1.0`) est rendue côté serveur. Tout est relatif à l'utilisateur connecté (`/me`). Chaque outil porte un tag `lecture` ou `ecriture` (le serveur retire `ecriture` en mode lecture seule) ; `envoi` marque ce qui envoie un courrier, une invitation ou une réponse de réunion.

## 1. Outils

| Outil | Méthode | Chemin | Portée déléguée |
|---|---|---|---|
| `outlook_list_mail_folders` | GET | `/me/mailFolders` | Mail.Read (ou ReadWrite) |
| `outlook_list_child_folders` | GET | `/me/mailFolders/{folderId}/childFolders` | Mail.Read (ou ReadWrite) |
| `outlook_get_mail_folder` | GET | `/me/mailFolders/{folderId}` | Mail.Read (ou ReadWrite) |
| `outlook_list_folder_messages` | GET | `/me/mailFolders/{folderId}/messages` | Mail.Read (ou ReadWrite) |
| `outlook_create_mail_folder` | POST | `/me/mailFolders` | Mail.ReadWrite |
| `outlook_create_child_folder` | POST | `/me/mailFolders/{folderId}/childFolders` | Mail.ReadWrite |
| `outlook_rename_mail_folder` | PATCH | `/me/mailFolders/{folderId}` | Mail.ReadWrite |
| `outlook_delete_mail_folder` | DELETE | `/me/mailFolders/{folderId}` | Mail.ReadWrite |
| `outlook_move_message` | POST | `/me/messages/{messageId}/move` | Mail.ReadWrite |
| `outlook_copy_message` | POST | `/me/messages/{messageId}/copy` | Mail.ReadWrite |
| `outlook_folder_messages_delta` | GET | `/me/mailFolders/{folderId}/messages/delta` | Mail.Read (ou ReadWrite) |
| `outlook_list_messages` | GET | `/me/messages` | Mail.Read (ou ReadWrite) |
| `outlook_search_messages` | GET | `/me/messages` | Mail.Read (ou ReadWrite) |
| `outlook_get_message` | GET | `/me/messages/{messageId}` | Mail.Read (ou ReadWrite) |
| `outlook_list_attachments` | GET | `/me/messages/{messageId}/attachments` | Mail.Read (ou ReadWrite) |
| `outlook_get_attachment` | GET | `/me/messages/{messageId}/attachments/{attachmentId}` | Mail.Read (ou ReadWrite) |
| `outlook_create_draft` | POST | `/me/messages` | Mail.ReadWrite |
| `outlook_update_message` | PATCH | `/me/messages/{messageId}` | Mail.ReadWrite |
| `outlook_send_draft` | POST | `/me/messages/{messageId}/send` | Mail.Send |
| `outlook_send_mail` | POST | `/me/sendMail` | Mail.Send |
| `outlook_reply_message` | POST | `/me/messages/{messageId}/reply` | Mail.Send |
| `outlook_reply_all_message` | POST | `/me/messages/{messageId}/replyAll` | Mail.Send |
| `outlook_forward_message` | POST | `/me/messages/{messageId}/forward` | Mail.Send |
| `outlook_delete_message` | DELETE | `/me/messages/{messageId}` | Mail.ReadWrite |
| `outlook_list_calendars` | GET | `/me/calendars` | Calendars.Read (ou ReadWrite) |
| `outlook_calendar_view` | GET | `/me/calendar/calendarView` | Calendars.Read (ou ReadWrite) |
| `outlook_list_events` | GET | `/me/events` | Calendars.Read (ou ReadWrite) |
| `outlook_get_event` | GET | `/me/events/{eventId}` | Calendars.Read (ou ReadWrite) |
| `outlook_create_event` | POST | `/me/events` | Calendars.ReadWrite |
| `outlook_update_event` | PATCH | `/me/events/{eventId}` | Calendars.ReadWrite |
| `outlook_delete_event` | DELETE | `/me/events/{eventId}` | Calendars.ReadWrite |
| `outlook_cancel_event` | POST | `/me/events/{eventId}/cancel` | Calendars.ReadWrite |
| `outlook_accept_event` | POST | `/me/events/{eventId}/accept` | Calendars.ReadWrite |
| `outlook_decline_event` | POST | `/me/events/{eventId}/decline` | Calendars.ReadWrite |
| `outlook_tentatively_accept_event` | POST | `/me/events/{eventId}/tentativelyAccept` | Calendars.ReadWrite |
| `outlook_find_meeting_times` | POST | `/me/findMeetingTimes` | Calendars.Read (ou ReadWrite) |
| `outlook_get_schedule` | POST | `/me/calendar/getSchedule` | Calendars.Read (ou ReadWrite) |
| `outlook_list_event_instances` | GET | `/me/events/{eventId}/instances` | Calendars.Read (ou ReadWrite) |
| `outlook_list_contacts` | GET | `/me/contacts` | Contacts.Read (ou ReadWrite) |
| `outlook_get_contact` | GET | `/me/contacts/{contactId}` | Contacts.Read (ou ReadWrite) |
| `outlook_create_contact` | POST | `/me/contacts` | Contacts.ReadWrite |
| `outlook_update_contact` | PATCH | `/me/contacts/{contactId}` | Contacts.ReadWrite |
| `outlook_delete_contact` | DELETE | `/me/contacts/{contactId}` | Contacts.ReadWrite |
| `outlook_list_contact_folders` | GET | `/me/contactFolders` | Contacts.Read (ou ReadWrite) |
| `outlook_search_people` | GET | `/me/people` | People.Read |
| `outlook_get_me` | GET | `/me` | User.Read |
| `outlook_get_mailbox_settings` | GET | `/me/mailboxSettings` | MailboxSettings.Read |
| `outlook_list_master_categories` | GET | `/me/outlook/masterCategories` | MailboxSettings.Read |
| `outlook_list_message_rules` | GET | `/me/mailFolders/{folderId}/messageRules` | Mail.Read (ou ReadWrite) |

Notes de conception :
- Paramètres OData sans préfixe `$` (`select`, `filter`, `search`, `top`, `skip`, `orderby`, `expand`, `count`, `skiptoken`, `deltatoken`) ; `includeHiddenFolders`, `startDateTime`, `endDateTime` s'écrivent tels quels.
- Dossiers en première classe : `outlook_list_mail_folders` ne renvoie que la racine (Graph pagine à 10 par défaut : `top=100`) ; `outlook_list_child_folders` descend d'un niveau (récursion à la charge de l'agent via `childFolderCount`) ; les messages d'un sous-dossier se lisent par `outlook_list_folder_messages`. Les noms connus (`inbox`, `sentitems`, `drafts`, `deleteditems`, `archive`, `junkemail`, `outbox`, `msgfolderroot`) sont admis partout où un `folderId` est attendu.
- `search` et `filter` sont exclusifs sur les messages ; `search` ignore `orderby`. `outlook_search_messages` est la variante avec `search` obligatoire.
- Pagination : Graph renvoie un `@odata.nextLink` absolu que le manuel ne suit pas ; l'agent recopie `skip` (messages, contacts, événements) ou `skiptoken` (calendarView, instances, delta, recherche) dans l'appel suivant.
- Corps de message : argument `Prefer` (en-tête, `header_fields`) en `enum` `outlook.body-content-type="text"` ou `"html"` ; `text` recommandé. Calendrier : `Prefer` avec `outlook.timezone="Europe/Paris"` (motif contrôlé). `outlook_folder_messages_delta` accepte `Prefer` libre (`odata.maxpagesize=50`).
- Identifiants Graph : `^[A-Za-z0-9_=-]{1,512}$` (base64url avec `=`, `-`, `_`).
- Corps POST/PATCH : un seul objet `body_field` par outil ; `additionalProperties: true` seulement dans les objets de corps.
- `outlook_send_draft` n'a pas de corps Graph ; l'argument `body` (objet vide `{}`) n'existe que pour satisfaire le contrôle « POST sans body_field ».
- « Tout marquer lu » dans un dossier : aucune opération Graph ; lister puis `outlook_update_message` message par message.
- Ids qui changent : `outlook_move_message` renvoie un NOUVEL identifiant de message.

## 2. Jeton

Jeton d'accès **délégué** (compte de l'utilisateur, flux code d'autorisation + PKCE ou device code) obtenu hors manuel, fourni au serveur dans `OUTLOOK_GRAPH_TOKEN` ; l'auth est un `api_key` `Authorization: Bearer ${OUTLOOK_GRAPH_TOKEN}`. Un jeton dure environ une heure : le renouvellement (refresh token, scope `offline_access`) est du ressort du serveur, pas du manuel.

**Substitution partielle vérifiée** avec python-utcp 1.2.0 : la chaîne `Bearer ${OUTLOOK_GRAPH_TOKEN}` est bien substituée à l'intérieur (variable client `outlook_OUTLOOK_GRAPH_TOKEN`), le serveur de test loopback a reçu `Authorization: Bearer <valeur>`. Le repli `${OUTLOOK_GRAPH_AUTHORIZATION}` n'a donc pas été nécessaire.

Scopes à demander (délégués) :

| Scope | Outils |
|---|---|
| `User.Read` | `outlook_get_me` |
| `Mail.Read` (ou `Mail.ReadWrite`) | lecture de dossiers, messages, pièces jointes, delta, règles |
| `Mail.ReadWrite` | création, modification, déplacement, copie, suppression de messages et de dossiers, brouillons |
| `Mail.Send` | `outlook_send_mail`, `outlook_send_draft`, `outlook_reply_message`, `outlook_reply_all_message`, `outlook_forward_message` |
| `MailboxSettings.Read` | `outlook_get_mailbox_settings`, `outlook_list_master_categories`, `outlook_list_message_rules` (les règles demandent `MailboxSettings.Read`) |
| `Calendars.ReadWrite` | tous les outils calendrier (lecture : `Calendars.Read` suffit ; `outlook_get_schedule` : `Calendars.ReadBasic` suffit) |
| `Contacts.ReadWrite` | tous les outils contacts (lecture : `Contacts.Read` suffit) |
| `People.Read` | `outlook_search_people` |
| `offline_access` | renouvellement du jeton (hors manuel) |

Le tableau de la section 1 donne le scope minimal par outil ; envoi d'invitation, de réponse de réunion ou d'annulation : le scope calendrier suffit, l'envoi est fait par Exchange.

## 3. Non couvert et pourquoi

- **Version MIME** (`GET /me/messages/{id}/$value`) : le client python-utcp 1.2.0 traite tout `$nom` d'une URL comme une variable (`\$([a-zA-Z0-9_]+)`) et lève `UtcpVariableNotFound` (« outlook_value ») au chargement. Aucun échappement n'existe. Le chemin est donc omis. Valide pour le validateur du dépôt, mais inutilisable avec le SDK de référence.
- **Pièces jointes > 3 Mio** (`createUploadSession` puis envois par blocs `PUT` d'octets bruts avec `Content-Range`) : le template `http` n'envoie que du JSON. Les petites pièces (`fileAttachment` en base64) passent dans `attachments` des brouillons et de `sendMail`. Le téléchargement d'une pièce lourde (`contentBytes`) est lu en JSON, peut dépasser le contexte de l'agent.
- **Delta** : couvert pour les messages d'un dossier seulement ; omis pour l'arbre de dossiers (`/me/mailFolders/delta`), les événements (`/me/calendarView/delta`) et les contacts (`/me/contactFolders/{id}/contacts/delta`).
- **Boîtes partagées, délégation** (`/users/{id}/...`) : hors périmètre, le manuel vise `/me`. Demande `Mail.Read.Shared` / `Mail.ReadWrite.Shared` etc.
- **Calendriers autres que le défaut** : `outlook_calendar_view`, `outlook_list_events` et la création lisent/écrivent le calendrier par défaut ; `/me/calendars/{id}/...` n'est pas décliné (liste seule via `outlook_list_calendars`).
- **Autres omis (limite de 50 outils)** : `createReply`/`createForward` (brouillons de réponse), recherche de messages en `POST /search/query`, dossiers de contacts en écriture, signature/règles en écriture, `mailboxSettings` en PATCH, catégories en écriture, `snooze`/`flag` de dossier, éléments `todo`, Teams et réunions en ligne détaillées, `@mentions`, conversations, vues d'attachements `item`/`reference`.
- **Exchange sur site** (on-premises, EWS) : pas de Graph ; hors périmètre. Comptes personnels Microsoft : `People.Read`, `outlook_find_meeting_times` et `outlook_get_schedule` n'y sont pas tous disponibles (getSchedule : comptes professionnels ou scolaires seulement).
- **API People** : Microsoft la déclare en maintenance (recommande `POST /search/query` avec `entityTypes: person`), mais elle reste en service ; `outlook_search_people` l'utilise.
