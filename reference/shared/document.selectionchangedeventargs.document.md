
# <a name="documentselectionchangedeventargs.document-property"></a>Propriété DocumentSelectionChangedEventArgs.document
Obtient un objet **Document** qui représente le document ayant déclenché l’événement **SelectionChanged**.

|||
|:-----|:-----|
|**Hôtes :**|Excel, Word|
|**Ajouté dans**|1.1|




```js
var myDoc = eventArgsObj.document;
```


## <a name="return-value"></a>Valeur renvoyée

Objet [Document](../../reference/shared/document.md) qui représente le document ayant déclenché l’événement [SelectionChanged](../../reference/shared/document.selectionchanged.event.md).


## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Niveau d’autorisation minimal**|[Restreint](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## <a name="support-history"></a>Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.0|Introduit|
