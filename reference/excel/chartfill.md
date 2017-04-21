# <a name="chartfill-object-javascript-api-for-excel"></a>Objet ChartFill (API JavaScript pour Excel)

Représente le format de remplissage d’un élément de graphique.

## <a name="properties"></a>Propriétés

Aucun

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|Supprime la couleur de remplissage d’un élément de graphique.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[setSolidColor(color: string)](#setsolidcolorcolor-string)|void|Définit le format de remplissage d’un élément de graphique sur une couleur unie.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="clear"></a>Effacer
Supprime la couleur de remplissage d’un élément de graphique.

#### <a name="syntax"></a>Syntaxe
```js
chartFillObject.clear();
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

### <a name="setsolidcolorcolor-string"></a>setSolidColor(color: string)
Définit le format de remplissage d’un élément de graphique sur une couleur unie.

#### <a name="syntax"></a>Syntaxe
```js
chartFillObject.setSolidColor(color);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|color|string|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|

#### <a name="returns"></a>Renvoie
void

#### <a name="examples"></a>Exemples

Définir le rouge comme couleur d’arrière-plan de Chart1

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    

    chart.format.fill.setSolidColor("#FF0000");

    return ctx.sync().then(function() {
            console.log("Chart1 Background Color Changed.");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
