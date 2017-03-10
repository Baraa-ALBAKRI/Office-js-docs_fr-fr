# <a name="chartlineformat-object-javascript-api-for-excel"></a>Objet ChartLineFormat (API JavaScript pour Excel)

Regroupe les options de mise en forme pour les éléments de ligne.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|color|string|Code couleur HTML qui représente la couleur des lignes dans le graphique.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|Désactiver le format de ligne d’un élément de graphique.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="clear"></a>Effacer
Désactiver le format de ligne d’un élément de graphique.

#### <a name="syntax"></a>Syntaxe
```js
chartLineFormatObject.clear();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples

Désactiver le format des lignes de quadrillage principal sur l’axe des ordonnées du graphique « Chart1 »

```js
Excel.run(function (ctx) { 
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueAxis.majorGridlines;    
    gridlines.format.line.clear();
    return ctx.sync().then(function() {
            console.log("Chart Major Gridlines Format Cleared");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

Définir le rouge comme couleur des lignes de quadrillage principal sur l’axe des ordonnées.

```js
Excel.run(function (ctx) {
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueAxis.majorGridlines;
    gridlines.format.line.color = "#FF0000";
    return ctx.sync().then(function () {
        console.log("Chart Gridlines Color Updated");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
