# <a name="pivottable-object-javascript-api-for-excel"></a>Objet PivotTable (API JavaScript pour Excel)

Représente un tableau croisé dynamique Excel.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|name|chaîne|Nom du tableau croisé dynamique.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|feuille de calcul|[Worksheet](worksheet.md)|Feuille de calcul contenant le tableau croisé dynamique. En lecture seule.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[refresh()](#refresh)|void|Actualise le tableau croisé dynamique.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="refresh"></a>refresh()
Actualise le tableau croisé dynamique.

#### <a name="syntax"></a>Syntaxe
```js
pivotTableObject.refresh();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void
