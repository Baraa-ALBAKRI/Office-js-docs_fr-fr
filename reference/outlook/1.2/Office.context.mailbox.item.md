

# <a name="item"></a>élément

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). item

L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](Office.context.mailbox.item.md#itemtype-officemailboxenumsitemtype).

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Mode Outlook applicable| Composition ou lecture|

### <a name="example"></a>Exemple

L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a>Membres

#### <a name="attachments-arrayattachmentdetailssimpletypesmdattachmentdetails"></a>attachments :Array.<[AttachmentDetails](simple-types.md#attachmentdetails)>

Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.

##### <a name="type"></a>Type :

*   Array.<[AttachmentDetails](simple-types.md#attachmentdetails)>

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|

##### <a name="example"></a>Exemple

Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.

```JavaScript
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsrecipientsmd"></a>bcc :[Recipients](Recipients.md)

Obtient ou définit les destinataires en Cci (copie carbone invisible) d’un message. Mode composition uniquement.

##### <a name="type"></a>Type :

*   [Destinataires](Recipients.md)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition|

##### <a name="example"></a>Exemple

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodybodymd"></a>body :[Body](Body.md)

Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.

##### <a name="type"></a>Type :

*   [Corps](Body.md)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|
####  cc :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Destinataires](Recipients.md)

Obtient ou définit les destinataires en copie carbone (Cc) d’un message.

##### <a name="read-mode"></a>Mode lecture

La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.

##### <a name="compose-mode"></a>Mode composition

La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant de manipuler des destinataires sur la ligne **Cc** du message.

##### <a name="type"></a>Type :

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### <a name="example"></a>Exemple

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a>(nullable) conversationId :String

Obtient l’identificateur de la conversation qui contient un message particulier.

Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.

Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.

##### <a name="type"></a>Type :

*   Chaîne

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|
#### <a name="datetimecreated-date"></a>dateTimeCreated :Date

Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.

##### <a name="type"></a>Type :

*   Date

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|

##### <a name="example"></a>Exemple

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a>dateTimeModified :Date

Obtient la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.

##### <a name="type"></a>Type :

*   Date

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|

##### <a name="example"></a>Exemple

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimetimemd"></a>end :Date|[Time](Time.md)

Obtient ou définit la date et l’heure de fin du rendez-vous.

La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.

##### <a name="read-mode"></a>Mode lecture

La propriété `end` renvoie un objet `Date`.

##### <a name="compose-mode"></a>Mode composition

La propriété `end` renvoie un objet `Time`.

Quand vous utilisez la méthode [`Time.setAsync`](Time.md#setasync) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.

##### <a name="type"></a>Type :

*   Date | [Time](Time.md)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### <a name="example"></a>Exemple

L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](Time.md#setasync) de l’objet `Time`.

```JavaScript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailssimpletypesmdemailaddressdetails"></a>from :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.

Les propriétés `from` et [`sender`](Office.context.mailbox.item.md#sender) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.

##### <a name="type"></a>Type :

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|
#### <a name="internetmessageid-string"></a>internetMessageId :String

Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.

##### <a name="type"></a>Type :

*   Chaîne

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|

##### <a name="example"></a>Exemple

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a>itemClass :String

Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.

La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.

| Type | Description | Classe de l’élément |
| --- | --- | --- |
| Éléments de rendez-vous | Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`. | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| Éléments de message | Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base. | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.

##### <a name="type"></a>Type :

*   Chaîne

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|

##### <a name="example"></a>Exemple

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a>(nullable) itemId :String

Obtient l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.

L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange. La propriété `itemId` n’est pas identique à l’identificateur d’entrée Outlook.

La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](Office.context.mailbox.item.md#saveAsync) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](simple-types.md#asyncresult) dans la fonction de rappel.

##### <a name="type"></a>Type :

*   Chaîne

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|

##### <a name="example"></a>Exemple

Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypeofficemailboxenumsmditemtypestring"></a>itemType :[Office.MailboxEnums.ItemType](Office.MailboxEnums.md#itemtype-string)

Obtient le type d’élément représenté par une instance.

La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.

##### <a name="type"></a>Type :

*   [Office.MailboxEnums.ItemType](Office.MailboxEnums.md#itemtype-string)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### <a name="example"></a>Exemple

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationlocationmd"></a>location :String|[Location](Location.md)

Obtient ou définit le lieu d’un rendez-vous.

##### <a name="read-mode"></a>Mode lecture

La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.

##### <a name="compose-mode"></a>Mode composition

La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.

##### <a name="type"></a>Type :

*   String | [Location](Location.md)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### <a name="example"></a>Exemple

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a>normalizedSubject :String

Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.

La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](Office.context.mailbox.item.md#subject).

##### <a name="type"></a>Type :

*   Chaîne

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|

##### <a name="example"></a>Exemple

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailssimpletypesmdemailaddressdetailsrecipientsrecipientsmd"></a>optionalAttendees :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Recipients](Recipients.md)

Obtient ou définit la liste des adresses de messagerie des participants facultatifs.

##### <a name="read-mode"></a>Mode lecture

La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.

##### <a name="compose-mode"></a>Mode composition

La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes pour obtenir et définir les participants facultatifs d’une réunion.

##### <a name="type"></a>Type :

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### <a name="example"></a>Exemple

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailssimpletypesmdemailaddressdetails"></a>organizer :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.

##### <a name="type"></a>Type :

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|

##### <a name="example"></a>Exemple

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailssimpletypesmdemailaddressdetailsrecipientsrecipientsmd"></a>requiredAttendees :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Recipients](Recipients.md)

Obtient ou définit la liste des adresses de messagerie des participants obligatoires.

##### <a name="read-mode"></a>Mode lecture

La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.

##### <a name="compose-mode"></a>Mode composition

La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes pour obtenir et définir les participants requis à une réunion.

##### <a name="type"></a>Type :

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### <a name="example"></a>Exemple

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="resources-emailaddressdetailssimpletypesmdemailaddressdetails"></a>resources :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Obtient les ressources requises pour un rendez-vous. Mode lecture uniquement.

##### <a name="type"></a>Type :

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|
#### <a name="sender-emailaddressdetailssimpletypesmdemailaddressdetails"></a>sender :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.

Les propriétés [`from`](Office.context.mailbox.item.md#from) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.

##### <a name="type"></a>Type :

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|

##### <a name="example"></a>Exemple

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimetimemd"></a>start :Date|[Time](Time.md)

Obtient ou définit la date et l’heure de début du rendez-vous.

La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.

##### <a name="read-mode"></a>Mode lecture

La propriété `start` renvoie un objet `Date`.

##### <a name="compose-mode"></a>Mode composition

La propriété `start` renvoie un objet `Time`.

Quand vous utilisez la méthode [`Time.setAsync`](Time.md#setasync) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.

##### <a name="type"></a>Type :

*   Date | [Time](Time.md)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### <a name="example"></a>Exemple

L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](Time.md#setasync) de l’objet `Time`.

```JavaScript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectsubjectmd"></a>subject :String|[Subject](Subject.md)

Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.

La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.

##### <a name="read-mode"></a>Mode lecture

La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](Office.context.mailbox.item.md#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a>Mode composition

La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a>Type :

*   String | [Subject](Subject.md)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|
####  to :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Destinataires](Recipients.md)

Obtient ou définit les destinataires d’un message électronique.

##### <a name="read-mode"></a>Mode lecture

La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.

##### <a name="compose-mode"></a>Mode composition

La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant de manipuler des destinataires sur la ligne **À** du message.

##### <a name="type"></a>Type :

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### <a name="example"></a>Exemple

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a>Méthodes

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

Ajoute un fichier à un message ou un rendez-vous en pièce jointe.

La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.

L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](Office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`uri`| String||URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.|
|`attachmentName`| String||Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult). <br/>En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.<br/>En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.<br/><table class="nested-table"><thead><tr><th>Code d'erreur</th><th>Description</th></tr></thead><tbody><tr><td><code>AttachmentSizeExceeded</code></td><td>La pièce jointe dépasse la taille autorisée.</td></tr><tr><td><code>FileTypeNotSupported</code></td><td>La pièce jointe comporte une extension qui n’est pas autorisée.</td></tr><tr><td><code>NumberOfAttachmentsExceeded</code></td><td>Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</td></tr></tbody></table>|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Mode Outlook applicable| Composition|

##### <a name="example"></a>Exemple

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.

La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.

L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](Office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.

Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`itemId`| String||Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.|
|`attachmentName`| String||Objet de l’élément à joindre. La longueur maximale est de 255 caractères.|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult). <br/>En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.<br/>En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.<br/><table class="nested-table"><thead><tr><th>Code d'erreur</th><th>Description</th></tr></thead><tbody><tr><td><code>NumberOfAttachmentsExceeded</code></td><td>Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</td></tr></tbody></table>|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Mode Outlook applicable| Composition|

##### <a name="example"></a>Exemple

L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

#### <a name="displayreplyallformformdata"></a>displayReplyAllForm(formData)

Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.

Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.

Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.

Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Description|
|---|---|---|
|`formData`| String &#124; Object|Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.<br/>**OU**<br/>Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini comme suit :<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>htmlBody</code></td><td>String</td><td>&lt;optional&gt;</td><td>Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</td></tr><tr><td><code>attachments</code></td><td>Array.&lt;Object&gt;</td><td>&lt;optional&gt;</td><td>Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.<br/><br/><strong>Propriétés</strong><br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Description</th></tr></thead><tbody><tr><td><code>type</code></td><td>String</td><td>Indique le type de pièce jointe. Doit être <code>file</code> pour une pièce jointe de fichier ou <code>item</code> pour une pièce jointe d’élément.</td></tr><tr><td><code>name</code></td><td>String</td><td>Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</td></tr><tr><td><code>url</code></td><td>String</td><td>Utilisé uniquement si <code>type</code> est défini sur <code>file</code>. Il s’agit de l’URI de l’emplacement du fichier.</td></tr><tr><td><code>itemId</code></td><td>String</td><td>Utilisé uniquement si <code>type</code> est défini sur <code>item</code>. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</td></tr></tbody></table></td></tr><tr><td><code>callback</code></td><td>function</td><td>&lt;optional&gt;</td><td>Une fois la méthode exécutée, la fonction transmise au paramètre <code>callback</code> est appelée avec un seul paramètre, <code>asyncResult</code>, qui est un objet <a href="simple-types.md#asyncresult"><code>AsyncResult</code></a>. Pour plus d’informations, consultez la page relative à l’<a href="tutorial-asynchronous.html">utilisation de méthodes asynchrones</a>.</td></tr></tbody></table>|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|

##### <a name="examples"></a>Exemples

Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

Réponse avec un corps vide.

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

Réponse avec un corps.

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

Réponse avec un corps et la pièce jointe d’un fichier.

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

Réponse avec un corps et la pièce jointe d’un élément.

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a>displayReplyForm(formData)

Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.

Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.

Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.

Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Description|
|---|---|---|
|`formData`| String &#124; Object|Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.<br/>**OU**<br/>Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini comme suit :<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>htmlBody</code></td><td>String</td><td>&lt;optional&gt;</td><td>Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</td></tr><tr><td><code>attachments</code></td><td>Array.&lt;Object&gt;</td><td>&lt;optional&gt;</td><td>Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.<br/><br/><strong>Propriétés</strong><br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Description</th></tr></thead><tbody><tr><td><code>type</code></td><td>String</td><td>Indique le type de pièce jointe. Doit être <code>file</code> pour une pièce jointe de fichier ou <code>item</code> pour une pièce jointe d’élément.</td></tr><tr><td><code>name</code></td><td>String</td><td>Utilisé uniquement si <code>type</code> est défini sur <code>file</code>. Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</td></tr><tr><td><code>url</code></td><td>String</td><td>Utilisé uniquement si <code>type</code> est défini sur <code>file</code>. Il s’agit de l’URI de l’emplacement du fichier.</td></tr><tr><td><code>itemId</code></td><td>String</td><td>Utilisé uniquement si <code>type</code> est défini sur <code>item</code>. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</td></tr></tbody></table></td></tr><tr><td><code>callback</code></td><td>function</td><td>&lt;optional&gt;</td><td>Une fois la méthode exécutée, la fonction transmise au paramètre <code>callback</code> est appelée avec un seul paramètre, <code>asyncResult</code>, qui est un objet <a href="simple-types.md#asyncresult"><code>AsyncResult</code></a>. Pour plus d’informations, consultez la page relative à l’<a href="tutorial-asynchronous.html">utilisation de méthodes asynchrones</a>.</td></tr></tbody></table>|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|

##### <a name="examples"></a>Exemples

Le code suivant transmet une chaîne à la fonction `displayReplyForm`.

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

Réponse avec un corps vide.

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

Réponse avec un corps.

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

Réponse avec un corps et la pièce jointe d’un fichier.

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

Réponse avec un corps et la pièce jointe d’un élément.

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities-entitiessimpletypesmdentities"></a>getEntities() → {[Entities](simple-types.md#entities)}

Obtient les entités figurant dans l’élément sélectionné.

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|

##### <a name="returns"></a>Renvoie :

Type : 
[Entities](simple-types.md#entities)

##### <a name="example"></a>Exemple

L’exemple suivant accède aux entités des contacts dans l’élément actif.

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype-nullable-arraystringcontactsimpletypesmdcontactmeetingsuggestionsimpletypesmdmeetingsuggestionphonenumbersimpletypesmdphonenumbertasksuggestionsimpletypesmdtasksuggestion"></a>getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>}

Obtient un tableau de toutes les entités du type spécifié trouvées dans l’élément sélectionné.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Description|
|---|---|---|
|`entityType`| [Office.MailboxEnums.EntityType](Office.MailboxEnums.md#entitytype-string)|Une des valeurs d’énumération EntityType.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Mode Outlook applicable| Lecture|

##### <a name="returns"></a>Renvoie :

Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null. Si aucune entité du type spécifié n’est présente dans l’élément, la méthode renvoie un tableau vide. Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.

Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.

| Valeur de `entityType` | Type des objets du tableau renvoyé | Niveau d’autorisation requis |
| --- | --- | --- |
| `Address` | String | **Restricted** |
| `Contact` | Contact | **ReadItem** |
| `EmailAddress` | String | **ReadItem** |
| `MeetingSuggestion` | MeetingSuggestion | **ReadItem** |
| `PhoneNumber` | PhoneNumber | **Restricted** |
| `TaskSuggestion` | TaskSuggestion | **ReadItem** |
| `URL` | String | **Restricted** |

Type : 
Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>

##### <a name="example"></a>Exemple

L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans l’objet ou le corps de l’élément actif.

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename-nullable-arraystringcontactsimpletypesmdcontactmeetingsuggestionsimpletypesmdmeetingsuggestionphonenumbersimpletypesmdphonenumbertasksuggestionsimpletypesmdtasksuggestion"></a>getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>}

Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.

La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](https://msdn.microsoft.com/en-us/library/office/fp161166.aspx) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Description|
|---|---|---|
|`name`| String|Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|

##### <a name="returns"></a>Renvoie :

Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.

Type : 
Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>

#### <a name="getregexmatches-object"></a>getRegExMatches() → {Object}

Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.

La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.

Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](Body.md#getAsync) pour récupérer l’intégralité du corps de l’élément.

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|

##### <a name="returns"></a>Renvoie :

Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.

<dl class="param-type">

<dt>Type</dt>

<dd>Object</dd>

</dl>

##### <a name="example"></a>Exemple

L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename-nullable-arraystring"></a>getRegExMatchesByName(name) → (nullable) {Array.<String>}

Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.

La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.

Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Description|
|---|---|---|
|`name`| String|Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|

##### <a name="returns"></a>Renvoie :

Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.

<dl class="param-type">

<dt>Type</dt>

<dd>Array.<String></dd>

</dl>

##### <a name="example"></a>Exemple

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback-string"></a>getSelectedDataAsync(coercionType, [options], callback) → {String}

Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.

Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](Office.md#coerciontype-string)||Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.|
|`options`| Object| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction||Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult).

Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`. Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Mode Outlook applicable| Composition|

##### <a name="returns"></a>Renvoie :

Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.

<dl class="param-type">

<dt>Type</dt>

<dd>String</dd>

</dl>

##### <a name="example"></a>Exemple

```JavaScript
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.

Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`callback`| function||Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult).

Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](CustomProperties.md) dans la propriété `asyncResult.value`. Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.| 
|`userContext`| Objet| &lt;facultatif&gt;|Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel. Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### <a name="example"></a>Exemple

L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.

```JavaScript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a>removeAttachmentAsync(attachmentId, [options], [callback])

Supprime une pièce jointe d’un message ou d’un rendez-vous.

La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`attachmentId`| String||Identificateur de la pièce jointe à supprimer. La longueur maximale de la chaîne est de 100 caractères.|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult). <br/>En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.<br/><table class="nested-table"><thead><tr><th>Code d'erreur</th><th>Description</th></tr></thead><tbody><tr><td><code>InvalidAttachmentId</code></td><td>L’identificateur de la pièce jointe n’existe pas.</td></tr></tbody></table>|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Mode Outlook applicable| Composition|

##### <a name="example"></a>Exemple

Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="setselecteddataasyncdata-options-callback"></a>setSelectedDataAsync(data, [options], callback)

Insère les données dans le corps ou l’objet d’un message de manière asynchrone.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`data`| String||Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;optional&gt;</td><td>Si <code>text</code>, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</td></tr></tbody></table><p>Avec <code>html</code> et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur <code>InvalidDataFormat</code> est renvoyée.</p><p>Si la propriété <code>coercionType</code> n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.|</p>|
|`callback`| function||Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult). |

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.2|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Mode Outlook applicable| Composition|

##### <a name="example"></a>Exemple

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
