# <a name="charttitle-object-javascript-api-for-excel"></a>Objet ChartTitle (API JavaScript pour Excel)

Représente un objet de titre pour un graphique.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|overlay|bool|Valeur booléenne indiquant si le titre du graphique recouvre le graphique ou non.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|text|string|Représente le texte du titre d’un graphique.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|Valeur booléenne qui représente la visibilité d’un objet de titre de graphique.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartTitleFormat](charttitleformat.md)|Représente le format du titre d’un graphique, à savoir le format de remplissage et de la police. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes
Aucun


## <a name="method-details"></a>Détails des méthodes

### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

Obtenir la valeur `text` du titre du graphique Chart1.

```js
Excel.run(function (ctx) { 
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    

var title = chart.title;
title.load('text');
return ctx.sync().then(function() {
        console.log(title.text);
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
});
```

Définir la valeur `text` du titre du graphique sur « My Chart » et placer le titre en haut du graphique, sans superposition.

```js
Excel.run(function (ctx) { 
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    

chart.title.text= "My Chart"; 
chart.title.visible=true;
chart.title.overlay=true;

return ctx.sync().then(function() {
        console.log("Char Title Changed");
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
});
```
