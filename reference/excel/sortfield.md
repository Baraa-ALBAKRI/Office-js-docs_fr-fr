# <a name="sortfield-object-javascript-api-for-excel"></a>Objet SortField (API JavaScript pour Excel)

Représente une condition dans une opération de tri.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|ascending|bool|Indique si le tri s’effectue dans l’ordre croissant.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|Couleur ciblée par la condition si le tri est appliqué à la couleur ou à la police de la cellule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|dataOption|string|Options de tri supplémentaires pour ce champ. Les valeurs possibles sont les suivantes : Normal, TextAsNumber.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|Key|int|Colonne (ou ligne, selon l’orientation du tri) ciblée par la condition. Représentée sous forme d’un décalage par rapport à la première colonne (ou ligne).|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|sortOn|chaîne|Type de tri de cette condition. Les valeurs possibles sont les suivantes : Value, CellColor, FontColor, Icon.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|icône|[Icon](icon.md)|Représente l’icône ciblée par la condition si le tri est appliqué à l’icône de la cellule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes
Aucun

