# <a name="chartcollection-object-javascript-api-for-excel"></a>Objet ChartCollection (API JavaScript pour Excel)

Collection de tous les objets de graphique d’une feuille de calcul.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|count|int|Renvoie le nombre de graphiques dans la feuille de calcul. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|éléments|[Chart[]](chart.md)|Collection d’objets de graphique. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[add(type: string, sourceData: Range, seriesBy: string)](#addtype-string-sourcedata-range-seriesby-string)|[Chart](chart.md)|Crée un graphique.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Renvoie le nombre de graphiques dans la feuille de calcul.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[Chart](chart.md)|Extrait un graphique à l’aide de son nom. Si plusieurs graphiques portent le même nom, c’est le premier d’entre eux qui est renvoyé.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Chart](chart.md)|Extrait un graphique en fonction de sa position dans la collection.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[Chart](chart.md)|Extrait un graphique à l’aide de son nom. Si plusieurs graphiques portent le même nom, c’est le premier d’entre eux qui est renvoyé.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="addtype-string-sourcedata-range-seriesby-string"></a>add(type: string, sourceData: Range, seriesBy: string)
Crée un graphique.

#### <a name="syntax"></a>Syntaxe
```js
chartCollectionObject.add(type, sourceData, seriesBy);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|type|string|Représente le type d’un graphique. Les valeurs possibles sont les suivantes : ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc.|
|sourceData|Range|Objet Range correspondant aux données source.|
|seriesBy|string|Facultatif. Spécifie la façon dont les colonnes ou les lignes sont utilisées comme séries de données sur le graphique. Les valeurs possibles sont les suivantes : Auto (automatique), Columns (colonnes), Rows (lignes)|

#### <a name="returns"></a>Retourne
[Chart](chart.md)

#### <a name="examples"></a>Exemples

Ajouter un graphique dont la valeur `chartType` est « ColumnClustered » dans la feuille de calcul « Charts », avec la propriété `sourceData` définie sur la plage « A1:B4 » et la propriété `seriresBy` définie sur « auto ».

```js
Excel.run(function (ctx) { 
    var rangeSelection = "A1:B4";
    var range = ctx.workbook.worksheets.getItem(sheetName)
        .getRange(rangeSelection);
    var chart = ctx.workbook.worksheets.getItem(sheetName)
        .charts.add("ColumnClustered", range, "auto");    return ctx.sync().then(function() {
            console.log("New Chart Added");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcount"></a>getCount()
Renvoie le nombre de graphiques dans la feuille de calcul.

#### <a name="syntax"></a>Syntaxe
```js
chartCollectionObject.getCount();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Renvoie
int

### <a name="getitemname-string"></a>getItem(name: string)
Extrait un graphique à l’aide de son nom. Si plusieurs graphiques portent le même nom, c’est le premier d’entre eux qui est renvoyé.

#### <a name="syntax"></a>Syntaxe
```js
chartCollectionObject.getItem(name);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|name|string|Nom du graphique à extraire.|

#### <a name="returns"></a>Retourne
[Chart](chart.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var chartname = 'Chart1';
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartname);
    return ctx.sync().then(function() {
            console.log(chart.height);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var chartId = 'SamplChartId';
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartId);
    return ctx.sync().then(function() {
            console.log(chart.height);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```



#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
    return ctx.sync().then(function() {
            console.log(chart.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitematindex-number"></a>getItemAt(index: number)
Extrait un graphique en fonction de sa position dans la collection.

#### <a name="syntax"></a>Syntaxe
```js
chartCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[Chart](chart.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
    return ctx.sync().then(function() {
            console.log(chart.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitemornullobjectname-string"></a>getItemOrNullObject(name: string)
Extrait un graphique à l’aide de son nom. Si plusieurs graphiques portent le même nom, c’est le premier d’entre eux qui est renvoyé.

#### <a name="syntax"></a>Syntaxe
```js
chartCollectionObject.getItemOrNullObject(name);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|name|string|Nom du graphique à extraire.|

#### <a name="returns"></a>Retourne
[Chart](chart.md)
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

```js
Excel.run(function (ctx) { 
    var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
    charts.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < charts.items.length; i++)
        {
            console.log(charts.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Obtenir le nombre de graphiques

```js
Excel.run(function (ctx) { 
    var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
    charts.load('count');
    return ctx.sync().then(function() {
        console.log("charts: Count= " + charts.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

