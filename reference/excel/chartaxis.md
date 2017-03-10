# <a name="chartaxis-object-javascript-api-for-excel"></a>Objet ChartAxis (API JavaScript pour Excel)

Représente un axe unique dans un graphique.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|majorUnit|object|Représente l’intervalle entre deux graduations principales. Peut être défini sur une valeur numérique ou une chaîne vide.  La valeur renvoyée est toujours un nombre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|maximum|object|Représente la valeur maximale sur l’axe des ordonnées.  Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique).  La valeur renvoyée est toujours un nombre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|minimum|object|Représente la valeur minimale sur l’axe des ordonnées. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique). La valeur renvoyée est toujours un nombre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|minorUnit|object|Représente l’intervalle entre deux graduations secondaires. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique). La valeur renvoyée est toujours un nombre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartAxisFormat](chartaxisformat.md)|Représente la mise en forme d’un objet de graphique, à savoir le format des lignes et de la police. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|majorGridlines|[ChartGridlines](chartgridlines.md)|Renvoie un objet de quadrillage qui représente les lignes de quadrillage principal de l’axe spécifié. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|minorGridlines|[ChartGridlines](chartgridlines.md)|Renvoie un objet de quadrillage qui représente les lignes de quadrillage secondaire de l’axe spécifié. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|title|[ChartAxisTitle](chartaxistitle.md)|Représente le titre de l’axe. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes
Aucun


## <a name="method-details"></a>Détails des méthodes

### <a name="property-access-examples"></a>Exemples d’accès aux propriétés
Obtenir la valeur `maximum` de l’axe du graphique Chart1

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    var axis = chart.axes.valueAxis;
    axis.load('maximum');
    return ctx.sync().then(function() {
            console.log(axis.maximum);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Définir les valeurs `maximum`, `minimum`, `majorunit` et `minorunit` pour valueaxis. 

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    chart.axes.valueAxis.maximum = 5;
    chart.axes.valueAxis.minimum = 0;
    chart.axes.valueAxis.majorUnit = 1;
    chart.axes.valueAxis.minorUnit = 0.2;
    return ctx.sync().then(function() {
            console.log("Axis Settings Changed");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
