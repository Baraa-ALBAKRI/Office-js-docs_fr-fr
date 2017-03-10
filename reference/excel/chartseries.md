# <a name="chartseries-object-javascript-api-for-excel"></a>Objet ChartSeries (API JavaScript pour Excel)

Représente une série dans un graphique.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|name|string|Représente le nom d’une série dans un graphique.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartSeriesFormat](chartseriesformat.md)|Représente le format d’une série de graphique, à savoir le format de remplissage et des lignes. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|points|[ChartPointsCollection](chartpointscollection.md)|Représente la collection de tous les points de la série. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes
Aucun


## <a name="method-details"></a>Détails des méthodes

### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

Renommer la 1ère série de Chart1 sur « New Series Name »

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    chart.series.getItemAt(0).name = "New Series Name";
    return ctx.sync().then(function() {
            console.log("Series1 Renamed");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
