
# <a name="error.message-property"></a>Propriété error.message
Obtient une description détaillée de l’erreur.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Dernière modification dans la sélection**|1.1|

```
var errMessage = asyncResult.error.message;
```


## <a name="return-value"></a>Valeur renvoyée

Description de l’erreur sous forme de **chaîne**.


## <a name="remarks"></a>Remarques

L’objet **Error** et ses propriétés sont accessibles à partir de l’objet [AsyncResult](../../reference/shared/asyncresult.md) renvoyé dans la fonction transmise en tant qu’argument _callback_ d’une opération de données asynchrone.


## <a name="example"></a>Exemple

Pour déclencher une erreur, sélectionnez un tableau ou une matrice, puis appelez la fonction `setText`.


```js
function setText() {
    Office.context.document.setSelectedDataAsync("Hello World!",
        function (asyncResult) {
            if (asyncResult.status === "failed")
                var error = asyncResult.error;
            write(error.name + ": " + error.message);
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).

||**Office pour bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|**OWA pour les appareils**|**Outlook pour Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||v||||
|**Excel**|v|v|v|||
|**Outlook**|v|v||v|v|
|**PowerPoint**|v|v|v|||
|**Project**|v|||||
|**Word**|v|v|v|||

|||
|:-----|:-----|
|**Niveau d’autorisation minimal**|[Restreint](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|De contenu, de volet de tâche, Outlook|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## <a name="support-history"></a>Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de PowerPoint Online.|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire des compléments de contenu pour Access.|
|1.0|Introduit|
