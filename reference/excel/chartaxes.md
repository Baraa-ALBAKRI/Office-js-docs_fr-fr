# <a name="chartaxes-object-javascript-api-for-excel"></a>Objet ChartAxes (interface API JavaScript pour Excel)

Représente les axes du graphique.

## <a name="properties"></a>Propriétés

Aucun

## <a name="relationships"></a>Relations
| Relation | Type   |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|categoryAxis|[ChartAxis](chartaxis.md)|Représente l’axe des abscisses d’un graphique. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|seriesAxis|[ChartAxis](chartaxis.md)|Représente l’axe des séries d’un graphique 3D. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|valueAxis|[ChartAxis](chartaxis.md)|Représente l’axe des ordonnées. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="loadparam-object"></a>load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### <a name="syntax"></a>Syntaxe
```js
object.load(param);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### <a name="returns"></a>Retourne
void
