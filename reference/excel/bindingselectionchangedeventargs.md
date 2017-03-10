# <a name="bindingselectionchangedeventargs-object-javascript-api-for-excel"></a>Objet BindingSelectionChangedEventArgs (API JavaScript pour Excel)

Fournit des informations sur la liaison qui a déclenché l’événement SelectionChanged.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|columnCount|int|Obtient le nombre de colonnes sélectionnées.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|int|Obtient le nombre de lignes sélectionnées.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|startColumn|int|Obtient l’index de la première colonne de la sélection (de base zéro).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|startRow|int|Obtient l’index de la première ligne de la sélection (de base zéro).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|binding|[Binding](binding.md)|Obtient un objet Binding qui représente la liaison ayant déclenché l’événement SelectionChanged.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes
Aucun

