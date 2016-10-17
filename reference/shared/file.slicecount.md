
# <a name="file.slicecount-property"></a>Propriété File.sliceCount
Obtient le nombre de sections du fichier divisé.

|||
|:-----|:-----|
|**Hôtes :**|PowerPoint, Word|
|**Ajouté dans**|1.1|

```
var slices = file.sliceCount;
```


## <a name="return-value"></a>Valeur renvoyée

Nombre de sections :


## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|||||
|:-----|:-----|:-----|:-----|
||Office pour Bureau Windows|Office Online (dans un navigateur)|Office pour iPad|
|**PowerPoint**|v|v|v|
|**Word**|v|v|v|

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
|1.1|Prise en charge supplémentaire de PowerPoint et Word dans Office pour iPad.|
|1.0|Introduit|
