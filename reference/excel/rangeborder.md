# <a name="rangeborder-object-javascript-api-for-excel"></a>Objet RangeBorder (API JavaScript pour Excel)

Représente la bordure d’un objet.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|color|string|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|Représente l’identificateur de la bordure. En lecture seule. Les valeurs possibles sont les suivantes : EdgeTop (bord supérieur), EdgeBottom (bord inférieur), EdgeLeft (bord gauche), EdgeRight (bord droit), InsideVertical (intérieur vertical), InsideHorizontal (intérieur horizontal), DiagonalDown (diagonale vers le bas), DiagonalUp (diagonale vers le haut).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|sideIndex|string|Valeur constante qui indique un côté spécifique de la bordure. En lecture seule. Les valeurs possibles sont les suivantes : EdgeTop (bord supérieur), EdgeBottom (bord inférieur), EdgeLeft (bord gauche), EdgeRight (bord droit), InsideVertical (intérieur vertical), InsideHorizontal (intérieur horizontal), DiagonalDown (diagonale vers le bas), DiagonalUp (diagonale vers le haut).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|style|string|L’une des constantes de style de ligne déterminant le style de ligne de la bordure. Les valeurs possibles sont les suivantes : None (aucune), Continuous (continue), Dash (tirets), DashDot (ligne tiret-point), DashDotDot (ligne tiret-point-point), Dot (points), Double (double), SlantDashDot (ligne tiret-point oblique).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|weight|string|Spécifie l’épaisseur de la bordure entourant une plage. Les valeurs possibles sont les suivantes : Hairline (très fine), Thin (fine), Medium (moyenne), Thick (épaisse).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var borders = range.format.borders;
    borders.load('items');
    return ctx.sync().then(function() {
        console.log(borders.count);
        for (var i = 0; i < borders.items.length; i++)
        {
            console.log(borders.items[i].sideIndex);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
L’exemple suivant ajoute des bordures de grille autour de la plage.

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.borders.getItem('InsideHorizontal').style = 'Continuous';
    range.format.borders.getItem('InsideVertical').style = 'Continuous';
    range.format.borders.getItem('EdgeBottom').style = 'Continuous';
    range.format.borders.getItem('EdgeLeft').style = 'Continuous';
    range.format.borders.getItem('EdgeRight').style = 'Continuous';
    range.format.borders.getItem('EdgeTop').style = 'Continuous';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

