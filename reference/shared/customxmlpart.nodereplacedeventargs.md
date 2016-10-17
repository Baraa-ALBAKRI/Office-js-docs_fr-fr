
# <a name="nodereplacedeventargs-object"></a>NodeReplacedEventArgs, objet
Fournit des informations sur le nœud remplacé qui a déclenché l’événement [dataNodeReplaced](../../reference/shared/customxmlpart.datanodereplaced.event.md).

|||
|:-----|:-----|
|**Hôtes :**|Word|
|**Disponible dans l’[ensemble de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Dernière modification dans**|1.1|

```
NodeReplacedEventArgs
```


## <a name="members"></a>Membres


**Propriétés**


|**Nom**|**Description**|
|:-----|:-----|
|[isUndoRedo](../../reference/shared/customxmlpart.isundoredo.md)|Obtient des informations indiquant si le nœud remplacé a été inséré dans le cadre d’une opération d’annulation ou de rétablissement effectuée par l’utilisateur.|
|[newNode](../../reference/shared/customxmlpart.newnode.md)|Obtient le nouveau nœud.|
|[oldNode](../../reference/shared/customxmlpart.oldnode.md)|Obtient l’ancien nœud (remplacé).|

## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cet objet est pris en charge dans l’application hôte Office correspondante. Une cellule vide indique que l’application hôte Office ne prend pas en charge cet objet.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office pour bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
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
