# <a name="rangefont-object-javascript-api-for-excel"></a>Objet RangeFont (API JavaScript pour Excel)

Cet objet représente les attributs de police (nom de la police, taille de police, couleur, etc.) d’un objet.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|bold|bool|Représente le format de police Gras.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|Représentation sous forme de code couleur HTML de la couleur du texte. Par exemple, #FF0000 représente le rouge.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|italic|bool|Représente le format de police Italique.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Nom de la police (par exemple « Calibri »)|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|size|Double|Taille de police|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|underline|string|Type de soulignement appliqué à la police. Les valeurs possibles sont les suivantes : None (aucun), Single (simple), Double (double) SingleAccountant (comptable simple), DoubleAccountant (comptable double).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes
Aucun


## <a name="method-details"></a>Détails des méthodes

### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var rangeFont = range.format.font;
    rangeFont.load('name');
    return ctx.sync().then(function() {
        console.log(rangeFont.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
L’exemple ci-dessous définit le nom de la police. 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.font.name = 'Times New Roman';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```