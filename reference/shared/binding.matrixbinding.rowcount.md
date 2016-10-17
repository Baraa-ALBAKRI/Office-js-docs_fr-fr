
# <a name="matrixbinding.rowcount-property"></a>Propriété MatrixBinding.rowCount
Obtient le nombre de lignes de la structure de données de la matrice, sous forme de valeur entière.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, PowerPoint, Project, Word|
|**Disponible dans l’[ensemble de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings|
|**Dernière modification dans la sélection**|1.1|

```
var rowCount = bindingObj.rowCount;
```


## <a name="return-value"></a>Valeur renvoyée

Nombre de lignes de l’objet [MatrixBinding](../../reference/shared/binding.matrixbinding.md) spécifié.


## <a name="example"></a>Exemple




```js
function showBindingRowCount() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        write("Rows: " + asyncResult.value.rowCount);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette propriété est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette propriété.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de conditions requises**|MatrixBindings|
|**Niveau d’autorisation minimal**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## <a name="support-history"></a>Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad.|
|1.0|Introduit|
