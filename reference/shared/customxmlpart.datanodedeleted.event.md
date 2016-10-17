
# <a name="customxmlpart.datanodedeleted-event"></a>Événement CustomXmlPart.dataNodeDeleted
Se produit quand un nœud est supprimé.

|||
|:-----|:-----|
|**Hôtes :**|Word|
|**Disponible dans l’[ensemble de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Dernière modification dans**|1.1|

```
Office.EventType.DataNodeDeleted
```


## <a name="remarks"></a>Remarques

Pour ajouter un gestionnaire d’événements pour l’événement **dataNodeDeleted**, utilisez la méthode [addHandlerAsync](../../reference/shared/customxmlpart.addhandlerasync.md) de l’objet **CustomXmlPart**.


## <a name="example"></a>Exemple




```js
function addNodeDeletedEvent() {
    Office.context.document.customXmlParts.getByIdAsync("{3BC85265-09D6-4205-B665-8EB239A8B9A1}", function (result) {
        var xmlPart = result.value;
        xmlPart.addHandlerAsync(Office.EventType.DataNodeDeleted, function (eventArgs) {
            write("A node has been deleted.");
        });
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```




## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cet événement est pris en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cet événement.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).

||**Office pour bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||||
|**Excel**||||
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de conditions requises**|CustomXmlParts|
|**Niveau d’autorisation minimal**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## <a name="support-history"></a>Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de Word dans Office pour iPad.|
|1.0|Introduit|
