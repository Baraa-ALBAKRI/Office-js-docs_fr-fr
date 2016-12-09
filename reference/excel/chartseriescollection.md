# <a name="chartseriescollection-object-javascript-api-for-excel"></a>Objet ChartSeriesCollection (interface API JavaScript pour Excel)

Représente une collection de séries de graphique.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|count|int|Renvoie le nombre de séries de la collection. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|éléments|[ChartSeries[]](chartseries.md)|Collection d’objets chartSeries. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[getItemAt(index: number)](#getitematindex-number)|[ChartSeries](chartseries.md)|Extrait une série en fonction de sa position dans la collection.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getitematindex-number"></a>getItemAt(index: number)
Extrait une série en fonction de sa position dans la collection.

#### <a name="syntax"></a>Syntaxe
```js
chartSeriesCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[ChartSeries](chartseries.md)

#### <a name="examples"></a>Exemples

Obtenir le nom de la première série de la collection.

```js
Excel.run(function (ctx) { 
    var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
    seriesCollection.load('items');
    return ctx.sync().then(function() {
        console.log(seriesCollection.items[0].name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


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
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés
Obtenir le nom des séries de la collection

```js
Excel.run(function (ctx) { 
    var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
    seriesCollection.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < seriesCollection.items.length; i++)
        {
            console.log(seriesCollection.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Obtenir le nombre de séries dans la collection

```js
Excel.run(function (ctx) { 
    var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
    seriesCollection.load('count');
    return ctx.sync().then(function() {
        console.log("series: Count= " + seriesCollection.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

