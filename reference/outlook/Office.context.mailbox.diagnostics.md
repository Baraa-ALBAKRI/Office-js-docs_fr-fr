

# <a name="diagnostics"></a>diagnostics

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics

Fournit des informations de diagnostic à un complément Outlook.

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

### <a name="members"></a>Membres

####  <a name="hostname-string"></a>hostName :String

Obtient une chaîne qui représente le nom de l’application hôte.

Chaîne qui peut avoir l’une des valeurs suivantes : `Outlook`, `Mac Outlook`, `OutlookIOS` ou `OutlookWebApp`.

##### <a name="type"></a>Type :

*   Chaîne

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|
####  <a name="hostversion-string"></a>hostVersion :String

Obtient une chaîne qui représente la version de l’application hôte ou du serveur Exchange Server.

Si le complément de messagerie s’exécute sur le client de bureau Outlook ou sur Outlook pour iOS, la propriété `hostVersion` renvoie la version de l’application hôte, Outlook. Dans Outlook Web App, la propriété renvoie la version du serveur Exchange. Exemple : la chaîne `15.0.468.0`.

##### <a name="type"></a>Type :

*   Chaîne

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|
####  <a name="owaview-string"></a>OWAView :String

Obtient une chaîne qui représente le mode d’affichage actuel dans Outlook Web App.

La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.

Si l’application hôte n’est pas Outlook Web App, l’accès à cette propriété génère la valeur `undefined`.

Outlook Web App a trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées :

*   `OneColumn`, qui est affiché lorsque l’écran est étroit. Outlook Web App offre une mise en page à une colonne sur l’ensemble de l’écran d’un smartphone.
*   `TwoColumns`, qui est affiché lorsque l’écran est plus large. Outlook Web App utilise ce mode sur la plupart des tablettes.
*   `ThreeColumns`, qui est affiché lorsque l’écran est large. Par exemple, Outlook Web App utilise ce mode dans une fenêtre en mode Plein écran sur un ordinateur de bureau.

##### <a name="type"></a>Type :

*   Chaîne

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.0|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|
