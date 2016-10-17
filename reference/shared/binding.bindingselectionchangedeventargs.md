
# <a name="bindingselectionchangedeventargs-object"></a>BindingSelectionChangedEventArgs, objet
Fournit des informations sur la liaison qui a déclenché l’événement [SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md).

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Dernière modification dans TableBinding**|1.1|

```
Office.EventType.BindingSelectionChanged
```


## <a name="members"></a>Membres


**Propriétés**


|**Nom**|**Description**|
|:-----|:-----|
|[binding](../../reference/shared/binding.bindingselectionchangedevent.binding.md)|Obtient un objet [Binding](../../reference/shared/binding.md) qui représente la liaison ayant déclenché l’événement **SelectionChanged**.|
|[columnCount](../../reference/shared/binding.bindingselectionchangedevent.columncount.md)|Obtient le nombre de colonnes sélectionnées.|
|[rowCount](../../reference/shared/binding.bindingselectionchangedevent.rowcount.md)|Obtient le nombre de lignes sélectionnées.|
|[startRow](../../reference/shared/binding.bindingselectionchangedevent.startrow.md)|Obtient l’index de la première ligne de la sélection (de base zéro).|
|[startColumn](../../reference/shared/binding.bindingselectionchangedevent.startcolumn.md)|Obtient l’index de la première colonne de la sélection (de base zéro).|
|[type](../../reference/shared/binding.bindingselectionchangedevent.type.md)|Obtient une valeur d’énumération [EventType](../../reference/shared/eventtype-enumeration.md) qui identifie le genre d’événement déclenché.|

## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## <a name="support-history"></a>Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire de la liaison de tableau dans les compléments pour Access.|
|1.0|Introduit|
