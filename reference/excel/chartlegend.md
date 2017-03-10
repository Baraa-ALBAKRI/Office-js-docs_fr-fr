# <a name="chartlegend-object-javascript-api-for-excel"></a>Objet ChartLegend (API JavaScript pour Excel)

Représente la légende d’un graphique.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|overlay|bool|Valeur booléenne indiquant si la légende du graphique doit chevaucher le corps principal du graphique.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|Position|string|Représente la position de la légende sur le graphique. Les valeurs possibles sont les suivantes : Top, Bottom, Left, Right, Corner, Custom.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|Valeur booléenne qui représente la visibilité d’un objet ChartLegend.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartLegendFormat](chartlegendformat.md)|Représente le format d’une légende de graphique, à savoir le format du remplissage et de la police. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes
Aucun


## <a name="method-details"></a>Détails des méthodes

### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

Obtenir la valeur `position` de la légende de graphique dans Chart1

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    var legend = chart.legend;
    legend.load('position');
    return ctx.sync().then(function() {
            console.log(legend.position);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Afficher la légende de Chart1 et la placer en haut du graphique

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    chart.legend.visible = true;
    chart.legend.position = "top"; 
    chart.legend.overlay = false; 
    return ctx.sync().then(function() {
            console.log("Legend Shown ");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
``` 
