

# <a name="settings.removehandlerasync-method"></a>Méthode Settings.removeHandlerAsync
Supprime un gestionnaire d’événements pour l’événement **settingsChanged**.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, PowerPoint|
|**Disponible dans l’[ensemble de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Paramètres|
|**Dernière modification dans**|1.1|

```js
Office.context.document.settings.removeHandlerAsync(eventType,  [,options], callback);
```


## <a name="parameters"></a>Paramètres



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _eventType_|[EventType](../../reference/shared/eventtype-enumeration.md)|Spécifie le type d’événement à supprimer. Obligatoire.||
| _options_|**objet**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants||
| _handler_|**chaîne**|Spécifie le nom du gestionnaire à supprimer. ||
| _asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
| _callback_|**objet**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## <a name="callback-value"></a>Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **removeHandlerAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Renvoie toujours **undefined** car il n’existe aucun objet ni aucune donnée à récupérer lors de la suppression d’un gestionnaire d’événements.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou **objet** défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|

## <a name="remarks"></a>Remarques

Si le paramètre facultatif _handler_ est omis lors de l’appel de la méthode **removeHandlerAsync**, tous les gestionnaires d’événements pour la valeur _eventType_ spécifiée sont supprimés.


## <a name="example"></a>Exemple




```js
function removeSettingsChangedEventHandler() {
    Office.context.document.settings.removeHandlerAsync(Office.EventType.SettingsChanged, MyHandler);
}

function MyHandler(eventArgs) {
    write('Event raised: ' + eventArgs.type);
    doSomethingWithSettings(eventArgs.settings);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office pour bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de conditions requises**|Paramètres|
|**Niveau d’autorisation minimal**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## <a name="support-history"></a>Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de PowerPoint.|
|1.1|Prise en charge supplémentaire de l’ajout et de la suppression des gestionnaires d’événements pour l’événement **SettingsChanged** dans les compléments de contenu pour Access.|
|1.0|Introduit|
