
# <a name="slice.index-property"></a>Propriété Slice.index
Obtient l’index de la section de fichier.

|||
|:-----|:-----|
|**Hôtes :**|PowerPoint, Word|
|**Disponible dans l’[ensemble de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Fichier|
|**Dernière modification dans**|1.1|

```
var sliceIndex = slice.index;
```


## <a name="return-value"></a>Valeur renvoyée

Index de la section de fichier.


## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette propriété est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette propriété.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office pour bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de conditions requises**|Fichier|
|**Niveau d’autorisation minimal**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## <a name="support-history"></a>Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de PowerPoint et Word dans Office pour iPad.|
|1.0|Introduit|
