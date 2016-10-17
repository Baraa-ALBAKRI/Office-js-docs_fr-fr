
# <a name="slice-object"></a>Slice, objet
Représente une section d’un fichier de document.

|||
|:-----|:-----|
|**Hôtes :**|PowerPoint, Word|
|**Disponible dans l’[ensemble de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Fichier|
|**Dernière modification dans**|1.1|

```
slice
```


## <a name="members"></a>Membres


**Propriétés**


|**Nom**|**Description**|
|:-----|:-----|
|**[data](../../reference/shared/slice.data.md)**|Obtient les données brutes de la section de fichier.|
|**[index](../../reference/shared/slice.index.md)**|Obtient l’index de la section de fichier.|
|**[size](../../reference/shared/slice.size.md)**|Obtient la taille de la section en octets.|

## <a name="remarks"></a>Remarques

L’accès à l’objet **Slice** s’effectue à l’aide de la méthode [File.getSliceAsync](../../reference/shared/file.getsliceasync.md).


## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cet objet est pris en charge dans l’application hôte Office correspondante. Une cellule vide indique que l’application hôte Office ne prend pas en charge cet objet.

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
