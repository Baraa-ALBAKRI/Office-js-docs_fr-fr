
# <a name="settings.get-method"></a>Méthode Settings.get
Récupère le paramètre spécifié.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, PowerPoint, Word|
|**Disponible dans l’[ensemble de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Paramètres|
|**Dernière modification dans**|1.1|

```js
var mySetting = Office.context.document.settings.get(name);
```


## <a name="parameters"></a>Paramètres



_name_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type : **chaîne**

&nbsp;&nbsp;&nbsp;&nbsp;Nom respectant l’emploi des majuscules et minuscules pour le paramètre à récupérer.

    



## <a name="return-value"></a>Valeur renvoyée

Un **objet** dont les noms de propriétés sont mappés à des valeurs sérialisées JSON.


## <a name="example"></a>Exemple




```js
function displayMySetting() {
    write('Current value for mySetting: ' + Office.context.document.settings.get('mySetting'));
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette propriété est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette propriété.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office pour bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de conditions requises**|Paramètres|
|**Niveau d’autorisation minimal**|[Restreint](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## <a name="support-history"></a>Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de PowerPoint Online.|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire de la création de paramètres dans les compléments de contenu pour Access.|
|1.0|Introduit|
