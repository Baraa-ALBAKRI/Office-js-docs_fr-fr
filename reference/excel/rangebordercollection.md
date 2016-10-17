# <a name="rangebordercollection-object-(javascript-api-for-excel)"></a>Objet RangeBorderCollection (interface API JavaScript pour Excel)

Représente les objets de bordure qui composent la bordure de la plage.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|count|int|Nombre d’objets de bordure de la collection. En lecture seule.|
|items|[RangeBorder[]](rangeborder.md)|Collection d’objets rangeBorder. En lecture seule.|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[getItem(index: string)](#getitemindex-string)|[RangeBorder](rangeborder.md)|Obtient un objet de bordure à l’aide de son nom.|
|[getItemAt(index: number)](#getitematindex-number)|[RangeBorder](rangeborder.md)|Obtient un objet de bordure à l’aide de son indice.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## <a name="method-details"></a>Détails des méthodes


### <a name="getitem(index:-string)"></a>getItem(index: string)
Obtient un objet de bordure à l’aide de son nom. 

#### <a name="syntax"></a>Syntaxe
```js
rangeBorderCollectionObject.getItem(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|string|Valeur d’indice de l’objet de bordure à récupérer.  Les valeurs possibles sont les suivantes : EdgeTop (bord supérieur), EdgeBottom (bord inférieur), EdgeLeft (bord gauche), EdgeRight (bord droit), InsideVertical (intérieur vertical), InsideHorizontal (intérieur horizontal), DiagonalDown (diagonale vers le bas), DiagonalUp (diagonale vers le haut).|

#### <a name="returns"></a>Retourne
[RangeBorder](rangeborder.md)

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var borderName = 'EdgeTop';
    var border = range.format.borders.getItem(borderName);
    border.load('style');
    return ctx.sync().then(function() {
            console.log(border.style);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


#### <a name="examples"></a>Exemples
```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var border = ctx.workbook.borders.getItemAt(0);
    border.load('sideIndex');
    return ctx.sync().then(function() {
            console.log(border.sideIndex);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
Obtient un objet de bordure à l’aide de son indice.

#### <a name="syntax"></a>Syntaxe
```js
rangeBorderCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[RangeBorder](rangeborder.md)

#### <a name="examples"></a>Exemples
```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var border = ctx.workbook.borders.getItemAt(0);
    border.load('sideIndex');
    return ctx.sync().then(function() {
            console.log(border.sideIndex);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="load(param:-object)"></a>load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### <a name="syntax"></a>Syntaxe
```js
object.load(param);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### <a name="returns"></a>Renvoie
void
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
L’exemple suivant ajoute une bordure de grille autour de la plage.

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
