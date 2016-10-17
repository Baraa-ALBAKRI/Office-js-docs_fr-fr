

# <a name="userprofile"></a>userProfile

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

### <a name="members"></a>Membres

####  <a name="displayname-:string"></a>displayName :String

Obtient le nom d’affichage de l’utilisateur.

##### <a name="type:"></a>Type :

*   Chaîne

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### <a name="example"></a>Exemple

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-:string"></a>emailAddress :String

Obtient l’adresse de messagerie SMTP de l’utilisateur.

##### <a name="type:"></a>Type :

*   Chaîne

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### <a name="example"></a>Exemple

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-:string"></a>timeZone :String

Obtient le fuseau horaire par défaut de l’utilisateur.

##### <a name="type:"></a>Type :

*   Chaîne

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### <a name="example"></a>Exemple

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```