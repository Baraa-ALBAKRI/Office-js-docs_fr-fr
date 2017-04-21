# <a name="chart-object-javascript-api-for-excel"></a>Objet Chart (API JavaScript pour Excel)

Représente un objet de graphique dans un classeur.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|height|Double|Représente la hauteur, exprimée en points, de l’objet de graphique.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|Extrait un graphique en fonction de sa position dans la collection. En lecture seule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|left|Double|Distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Représente le nom d’un objet de graphique.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|top|Double|Représente la distance, en points, entre le bord supérieur de l’objet et la partie supérieure de la ligne 1 (sur une feuille de calcul) ou le haut de la zone de graphique (sur un graphique).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|width|Double|Représente la largeur, en points, de l’objet de graphique.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|axes|[ChartAxes](chartaxes.md)|Représente les axes du graphique. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|dataLabels|[ChartDataLabels](chartdatalabels.md)|Représente les étiquettes des données sur le graphique. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|format|[ChartAreaFormat](chartareaformat.md)|Regroupe les propriétés de format de la zone de graphique. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|légende|[ChartLegend](chartlegend.md)|Représente la légende du graphique. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|série|[ChartSeriesCollection](chartseriescollection.md)|Représente une série ou une collection de séries dans le graphique. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|title|[ChartTitle](charttitle.md)|Représente le titre du graphique indiqué et comprend le texte, la visibilité, la position et la mise en forme du titre. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|feuille de calcul|[Worksheet](worksheet.md)|Feuille de calcul contenant le graphique actuel. En lecture seule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Supprime l’objet Graphique.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getImage(height: number, width: number, fittingMode: string)](#getimageheight-number-width-number-fittingmode-string)|[System.IO.Stream](system.io.stream.md)|Affiche le graphique sous forme d’image codée en Base64 ajustée aux dimensions spécifiées.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[setData(sourceData: Range, seriesBy: string)](#setdatasourcedata-range-seriesby-string)|void|Réinitialise les données source du graphique.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[setPosition(startCell: Range or string, endCell: Range or string)](#setpositionstartcell-range-or-string-endcell-range-or-string)|void|Positionne le graphique par rapport aux cellules dans la feuille de calcul.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="delete"></a>delete()
Supprime l’objet de graphique.

#### <a name="syntax"></a>Syntaxe
```js
chartObject.delete();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    chart.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getimageheight-number-width-number-fittingmode-string"></a>getImage(height: number, width: number, fittingMode: string)
Affiche le graphique sous forme d’image codée en Base64 ajusté aux dimensions spécifiées.

#### <a name="syntax"></a>Syntaxe
```js
chartObject.getImage(height, width, fittingMode);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|height|number|Facultatif. (Facultatif) Hauteur souhaitée de l’image produite.|
|width|number|Facultatif. (Facultatif) Largeur souhaitée de l’image produite.|
|fittingMode|string|Facultatif. (Facultatif) Méthode utilisée pour mettre à l’échelle le graphique aux dimensions spécifiées (si la hauteur et la largeur sont définies).  Les valeurs possibles sont les suivantes : Fit (ajuster), FitAndCenter (ajuster et centrer), Fill (remplir)|

#### <a name="returns"></a>Retourne
[System.IO.Stream](system.io.stream.md)

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    var image = chart.getImage();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```





### <a name="setdatasourcedata-range-seriesby-string"></a>setData(sourceData: Range, seriesBy: string)
Redéfinit les données sources du graphique.

#### <a name="syntax"></a>Syntaxe
```js
chartObject.setData(sourceData, seriesBy);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|sourceData|Range|Objet Range correspondant aux données source.|
|seriesBy|string|Facultatif. Spécifie la façon dont les colonnes ou les lignes sont utilisées comme séries de données sur le graphique. Les options disponibles sont : en mode automatique (par défaut), par lignes et par colonnes. Les valeurs possibles sont les suivantes : Auto (automatique), Columns (colonnes), Rows (lignes)|

#### <a name="returns"></a>Renvoie
void

#### <a name="examples"></a>Exemples

Définir `sourceData` sur « A1:B4 » et `seriesBy` sur « Columns »

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    var sourceData = "A1:B4";
    chart.setData(sourceData, "Columns");
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="setpositionstartcell-range-or-string-endcell-range-or-string"></a>setPosition(startCell: range ou string, endCell: range ou string)
Positionne le graphique par rapport aux cellules dans la feuille de calcul.

#### <a name="syntax"></a>Syntaxe
```js
chartObject.setPosition(startCell, endCell);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|startCell|range ou string|Cellule de début. Il s’agit de l’emplacement où le graphique sera déplacé. La cellule de début est la cellule supérieure gauche ou supérieure droite, selon les paramètres d’affichage droite-gauche définis par l’utilisateur.|
|endCell|range ou string|Facultatif. (Facultatif) Cellule de fin. Si une valeur est indiquée, la largeur et la hauteur du graphique seront définies de manière à couvrir entièrement cette cellule/plage.|

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples


```js
Excel.run(function (ctx) { 
    var sheetName = "Charts";
    var rangeSelection = "A1:B4";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeSelection);
    var sourceData = sheetName + "!" + "A1:B4";
    var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("pie", range, "auto");
    chart.width = 500;
    chart.height = 300;
    chart.setPosition("C2", null);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

Obtenir un graphique nommé « Chart1 »

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    chart.load('name');
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

Mettre à jour un graphique, y compris son nom, sa position et ses dimensions

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    chart.name="New Name";
    chart.top = 100;
    chart.left = 100;
    chart.height = 200;
    chart.width = 200;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Renommer le graphique, définir les dimensions du graphique sur 200 points en hauteur et en largeur. Déplacer Chart1 de 100 points vers le haut et vers la gauche. 

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
    chart.name="New Name";    
    chart.top = 100;
    chart.left = 100;
    chart.height =200;
    chart.width =200;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

