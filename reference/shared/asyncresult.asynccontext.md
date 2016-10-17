
# <a name="asyncresult.asynccontext-property"></a>Propriété AsyncResult.asyncContext
Obtient l’élément défini par l’utilisateur transmis au paramètre facultatif _asyncContext_ de la méthode appelée dans le même état que celui dans lequel il a été transmis.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Dernière modification dans**|1.1|

```
var myContext = asynchResult.asyncContext;
```


## <a name="return-value"></a>Valeur renvoyée

Renvoie l’élément défini par l’utilisateur (qui peut être de l’un des types JavaScript suivants : **chaîne**, **numérique**, **booléen**, **objet**, **tableau**, **null** ou **non défini**) transmis au paramètre _asyncContext_ facultatif de la méthode appelée. Renvoie **Undefined** si vous n’avez rien transmis au paramètre _asyncContext_.


## <a name="example"></a>Exemple




```js
function getDataWithContext() {
    var format = "Your data: ";
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, { asyncContext: format }, showDataWithContext);
}

 function showDataWithContext(asyncResult) {
    write(asyncResult.asyncContext + asyncResult.value);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```




## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|**OWA pour les appareils**|**Outlook pour Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**|v|||||
|**Excel**|v|v|v|||
|**Outlook**|v|v||v|v|
|**PowerPoint**|v|v|v|||
|**Project**||||||
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
|1.1|Prise en charge supplémentaire des compléments pour Access.|
|1.0|Introduit|
