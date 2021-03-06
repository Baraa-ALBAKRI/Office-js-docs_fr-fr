
# <a name="context.mailbox-property"></a>Propriété Context.mailbox
Obtient l’objet **mailbox** qui donne accès aux membres de l’API spécifiquement pour les compléments Outlook.

|||
|:-----|:-----|
|**Hôtes :**|Outlook|
|**Disponible dans l’[ensemble de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Boîte aux lettres|
|**Dernière modification dans**|1.0|

```js
var outlookOm = Office.context.mailbox;
```


## <a name="return-value"></a>Valeur renvoyée

Objet [mailbox](http://msdn.microsoft.com/library/a3880d3b-8a09-4cf9-9274-f2682cb3b769%28Office.15%29.aspx).


## <a name="example"></a>Exemple

La ligne de code suivante accède à l’objet [item](http://msdn.microsoft.com/library/ad288df1-3ca2-474c-bea4-c51f46e6fc43%28Office.15%29.aspx) de l’API JavaScript pour Office.


```js
// Access the Item object.
var item = Office.context.mailbox.item;

```




## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office pour bureau Windows**|**Office Online (dans un navigateur)**|**Outlook pour Mac**|
|:-----|:-----|:-----|:-----|
|**Outlook**|v|v|v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de conditions requises**|Boîte aux lettres|
|**Niveau d’autorisation minimal**|[Restreint](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Outlook|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## <a name="support-history"></a>Historique de prise en charge


|**Version**|**Modifications**|
|:-----|:-----|
|1.0|Introduit|
