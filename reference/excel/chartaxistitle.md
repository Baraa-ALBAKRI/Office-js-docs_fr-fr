# <a name="chartaxistitle-object-javascript-api-for-excel"></a>Objet ChartAxisTitle (API JavaScript pour Excel)

Représente le titre d’un axe de graphique.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|text|string|Représente le titre de l’axe.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|Valeur booléenne qui spécifie la visibilité d’un titre d’axe.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartAxisTitleFormat](chartaxistitleformat.md)|Représente le format du titre d’un axe de graphique. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes
Aucun


## <a name="method-details"></a>Détails des méthodes

### <a name="property-access-examples"></a>Exemples d’accès aux propriétés
Obtenir la valeur `text` du titre d’un axe de graphique à partir de l’axe des ordonnées de Chart1.

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    var title = chart.axes.valueAxis.title;
    title.load('text');
    return ctx.sync().then(function() {
            console.log(title.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Ajouter « Values » comme titre de l’axe des ordonnées

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    chart.axes.valueAxis.title.text = "Values";
    return ctx.sync().then(function() {
            console.log("Axis Title Added ");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
