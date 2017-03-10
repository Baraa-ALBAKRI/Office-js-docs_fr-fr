# <a name="chartfont-object-javascript-api-for-excel"></a>Objet ChartFont (API JavaScript pour Excel)

Cet objet représente les attributs de police (nom de police, taille de police, couleur, etc.) d’un objet de graphique.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|bold|bool|Représente le format de police Gras.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|Représentation sous forme de code couleur HTML de la couleur du texte. Par exemple, #FF0000 représente le rouge.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|italic|bool|Représente le format de police Italique.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Nom de la police (par exemple « Calibri »)|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|size|Double|Taille de la police (par exemple, 11)|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|underline|string|Type de soulignement appliqué à la police. Les valeurs possibles sont les suivantes : None, Single.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes
Aucun


## <a name="method-details"></a>Détails des méthodes

### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

Utiliser le titre du graphique comme exemple

```js
Excel.run(function (ctx) { 
    var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
    title.format.font.name = "Calibri";
    title.format.font.size = 12;
    title.format.font.color = "#FF0000";
    title.format.font.italic =  false;
    title.format.font.bold = true;
    title.format.font.underline = "None";
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Définir une police Calibri de taille 10, en gras et en rouge pour le titre du graphique. 

```js
Excel.run(function (ctx) { 
    var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
    title.format.font.name = "Calibri";
    title.format.font.size = 12;
    title.format.font.color = "#FF0000";
    title.format.font.italic =  false;
    title.format.font.bold = true;
    title.format.font.underline = "None";
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
