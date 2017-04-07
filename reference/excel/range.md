# <a name="range-object-javascript-api-for-excel"></a>Objet Range (API JavaScript pour Excel)

Une plage représente un ensemble constitué de cellules contiguës comme une cellule, une ligne, une colonne, un bloc de cellules, etc.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|adresse|string|Représente la référence de plage dans le style A1. La valeur d’adresse contient la référence de feuille (par exemple, Feuille1! A1:B4). En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|addressLocal|string|Représente la référence de la plage spécifiée dans le langage de l’utilisateur. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|cellCount|int|Nombre de cellules dans la plage. Cette API renvoie -1 si le nombre de cellules est supérieur à 2^31-1 (2 147 483 647). En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|columnCount|int|Représente le nombre total de colonnes dans la plage. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|columnHidden|bool|Indique si toutes les colonnes de la plage active sont masquées.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|columnIndex|int|Représente le numéro de colonne de la première cellule de la plage. Avec indice zéro. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulas|object[][]|Représente la formule dans le style de notation A1.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulasLocal|object[][]|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur. Par exemple, la formule « =SUM(A1, 1.5) » en anglais deviendrait « =SUMME(A1; 1,5) » en allemand.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulasR1C1|object[][]|Représente la formule dans le style de notation R1C1.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|hidden|bool|Indique si toutes les cellules de la plage active sont masquées. En lecture seule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|object[][]|Représente le code de format de nombre d’Excel pour une cellule donnée.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|int|Renvoie le nombre total de lignes de la plage. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowHidden|bool|Indique si toutes les lignes de la plage active sont masquées.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|rowIndex|int|Renvoie le numéro de ligne de la première cellule de la plage. Avec indice zéro. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|text|object[][]|Valeurs de texte de la plage spécifiée. La valeur de texte ne dépend pas de la largeur de la cellule. Le remplacement par le signe # qui se produit dans l’interface utilisateur d’Excel n’a aucun effet sur la valeur de texte renvoyée par l’API. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|valueTypes|string|Représente le type de données de chaque cellule. En lecture seule. Les valeurs possibles sont les suivantes : Unknown (inconnu), Empty (vide), String (chaîne), Integer (entier), Double (double), Boolean (valeur booléenne), Error (erreur).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Une cellule contenant une erreur renvoie la chaîne d’erreur.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|format|[RangeFormat](rangeformat.md)|Renvoie un objet de format, qui comprend les propriétés de police, de remplissage, de bordures, d’alignement, etc. de la plage. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|tri|[RangeSort](rangesort.md)|Représente le tri de plage de la plage actuelle. En lecture seule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|feuille de calcul|[Worksheet](worksheet.md)|Feuille de calcul contenant la plage. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[clear(applyTo: string)](#clearapplyto-string)|void|Supprime les valeurs et les propriétés de format, de remplissage, de bordure, etc. de la plage.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[delete(shift: string)](#deleteshift-string)|void|Supprime les cellules associées à la plage.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getBoundingRect(anotherRange: Range or string)](#getboundingrectanotherrange-range-or-string)|[Range](range.md)|Renvoie le plus petit objet de plage qui englobe les plages données. Par exemple, la valeur GetBoundingRect pour « B2:C5 » et « D10:E15 » est « B2:E16 ».|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Renvoie l’objet de plage qui contient une cellule donnée sur la base des numéros de ligne et de colonne. La cellule peut se trouver en dehors des limites de ses plages parent, pour peu qu’elle reste dans la grille de la feuille de calcul. L’emplacement de la cellule renvoyée est déterminé à partir de la cellule supérieure gauche de la plage.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumn(column: number)](#getcolumncolumn-number)|[Range](range.md)|Obtient une colonne contenue dans la plage.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumnsAfter(count: nombre)](#getcolumnsaftercount-number)|[Range](range.md)|Obtient un certain nombre de colonnes à droite de l’objet de plage actuel.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumnsBefore(count: number)](#getcolumnsbeforecount-number)|[Range](range.md)|Obtient un certain nombre de colonnes à gauche de l’objet de plage actuel.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getEntireColumn()](#getentirecolumn)|[Range](range.md)|Obtient un objet qui représente la colonne entière de la plage (par exemple, si la plage actuelle représente les cellules « B4:E11 », sa valeur `getEntireColumn` est une plage qui représente les colonnes « B:E »).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getEntireRow()](#getentirerow)|[Range](range.md)|Obtient un objet qui représente la ligne entière de la plage (par exemple, si la plage actuelle représente les cellules « B4:E11 », sa valeur `GetEntireRow` est une plage qui représente les lignes « 4:11 »).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getIntersection(anotherRange: Range or string)](#getintersectionanotherrange-range-or-string)|[Range](range.md)|Obtient l’objet de plage qui représente l’intersection rectangulaire des plages données.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getIntersectionOrNullObject(anotherRange: range ou string)](#getintersectionornullobjectanotherrange-range-or-string)|[Range](range.md)|Obtient l’objet de plage qui représente l’intersection rectangulaire des plages données. Si aucune intersection n’est trouvée, renvoie un objet Null.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastCell()](#getlastcell)|[Range](range.md)|Obtient la dernière cellule de la plage. Par exemple, la dernière cellule de la plage « B2:D5 » est « D5 ».|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastColumn()](#getlastcolumn)|[Range](range.md)|Obtient la dernière colonne de la plage. Par exemple, la dernière colonne de la plage « B2:D5 » est « D2:D5 ».|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastRow()](#getlastrow)|[Range](range.md)|Obtient la dernière ligne de la plage. Par exemple, la dernière ligne de la plage « B2:D5 » est « B5:D5 ».|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getOffsetRange(rowOffset: number, columnOffset: number)](#getoffsetrangerowoffset-number-columnoffset-number)|[Range](range.md)|Obtient un objet qui représente une plage décalée par rapport à la plage spécifiée. Les dimensions de la plage renvoyée correspondent à cette plage. Si la plage obtenue se retrouve en dehors des limites de grille de la feuille de calcul, une erreur est déclenchée.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getResizedRange(deltaRows: nombre, deltaColumns: nombre)](#getresizedrangedeltarows-number-deltacolumns-number)|[Range](range.md)|Obtient un objet de plage semblable à l’objet de plage actuel, mais avec le coin inférieur droit développé (ou contracté) selon un certain nombre de lignes et de colonnes.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRow(row: number)](#getrowrow-number)|[Range](range.md)|Obtient une ligne contenue dans la plage.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowsAbove(count: nombre)](#getrowsabovecount-number)|[Range](range.md)|Obtient un certain nombre de lignes au-dessus de l’objet de plage actuel.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowsBelow(count: number)](#getrowsbelowcount-number)|[Range](range.md)|Obtient un certain nombre de lignes en dessous de l’objet de plage actuel.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRange(valuesOnly: bool)](#getusedrangevaluesonly-bool)|[Range](range.md)|Renvoie la plage utilisée d’un objet de plage donné. Si aucune cellule n’est utilisée dans la plage, cette fonction génère une erreur ItemNotFound.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRangeOrNullObject(valuesOnly: bool)](#getusedrangeornullobjectvaluesonly-bool)|[Range](range.md)|Renvoie la plage utilisée d’un objet de plage donné. Si aucune cellule n’est utilisée dans la plage, cette fonction renvoie un objet null.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getVisibleView()](#getvisibleview)|[RangeView](rangeview.md)|Représente les lignes visibles de la plage en cours.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[insert(shift: string)](#insertshift-string)|[Range](range.md)|Insère une cellule ou une plage de cellules dans la feuille de calcul à la place d’une plage donnée et décale les autres cellules pour libérer de l’espace. Renvoie un nouvel objet Range dans l’espace vide qui s’est créé.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[merge(across: bool)](#mergeacross-bool)|void|Fusionne la plage de cellules dans une zone de la feuille de calcul.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[select()](#select)|void|Sélectionne la plage spécifiée dans l’interface utilisateur d’Excel.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[unmerge()](#unmerge)|void|Annule la fusion de la plage de cellules.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="clearapplyto-string"></a>clear(applyTo: string)
Supprime les valeurs et les propriétés de format, de remplissage, de bordure, etc. de la plage.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.clear(applyTo);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|applyTo|string|Facultatif. Détermine le type d’action de suppression. Les valeurs possibles sont les suivantes : `All` Option par défaut,`Formats`, ,`Contents` |

#### <a name="returns"></a>Renvoie
void

#### <a name="examples"></a>Exemples

L’exemple ci-dessous efface le format et le contenu de la plage. 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.clear();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="deleteshift-string"></a>delete(shift: string)
Supprime les cellules associées à la plage.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.delete(shift);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|Shift|string|Indique la façon dont les cellules doivent être décalées. Les valeurs possibles sont les suivantes : Up (vers le haut), Left (vers la gauche)|

#### <a name="returns"></a>Renvoie
void

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getboundingrectanotherrange-range-or-string"></a>getBoundingRect(anotherRange: Range or string)
Renvoie le plus petit objet de plage qui englobe les plages données. Par exemple, la valeur GetBoundingRect pour « B2:C5 » et « D10:E15 » est « B2:E16 ».

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getBoundingRect(anotherRange);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|anotherRange|range ou string|Nom, adresse ou objet de plage.|

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:G6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var range = range.getBoundingRect("G4:H8");
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // Prints Sheet1!D4:H8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcellrow-number-column-number"></a>getCell(row: number, column: number)
Renvoie l’objet de plage qui contient une cellule donnée sur la base des numéros de ligne et de colonne. La cellule peut se trouver en dehors des limites de ses plages parent, pour peu qu’elle reste dans la grille de la feuille de calcul. L’emplacement de la cellule renvoyée est déterminé à partir de la cellule supérieure gauche de la plage.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getCell(row, column);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|row|number|Numéro de ligne de la cellule à récupérer. Avec indice zéro.|
|column|number|Numéro de colonne de la cellule à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var cell = range.cell(0,0);
    cell.load('address');
    return ctx.sync().then(function() {
        console.log(cell.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcolumncolumn-number"></a>getColumn(column: number)
Obtient une colonne contenue dans la plage.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getColumn(column);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|column|number|Numéro de colonne de la plage à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet19";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getColumn(1);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!B1:B8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcolumnsaftercount-number"></a>getColumnsAfter(count: nombre)
Obtient un certain nombre de colonnes à droite de l’objet de plage actuel.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getColumnsAfter(count);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|count|number|Facultatif. Nombre de colonnes à inclure dans la plage obtenue. En règle générale, utilisez un nombre positif pour créer une plage en dehors de la plage actuelle. Vous pouvez également utiliser un nombre négatif pour créer une plage à l’intérieur de la plage actuelle. La valeur par défaut est 1.|

#### <a name="returns"></a>Retourne
[Range](range.md)

### <a name="getcolumnsbeforecount-number"></a>getColumnsBefore(count: nombre)
Obtient un certain nombre de colonnes à gauche de l’objet de plage actuel.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getColumnsBefore(count);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|count|number|Facultatif. Nombre de colonnes à inclure dans la plage obtenue. En règle générale, utilisez un nombre positif pour créer une plage en dehors de la plage actuelle. Vous pouvez également utiliser un nombre négatif pour créer une plage à l’intérieur de la plage actuelle. La valeur par défaut est 1.|

#### <a name="returns"></a>Retourne
[Range](range.md)

### <a name="getentirecolumn"></a>getEntireColumn()
Obtient un objet qui représente la colonne entière de la plage (par exemple, si la plage actuelle représente les cellules « B4:E11 », sa valeur `getEntireColumn` est une plage qui représente les colonnes « B:E »).

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getEntireColumn();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

Remarque : les propriétés de grille de la plage (valeurs, format de nombre, formules) contiennent la valeur `null`, car la plage en question est illimitée.

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeEC = range.getEntireColumn();
    rangeEC.load('address');
    return ctx.sync().then(function() {
        console.log(rangeEC.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getentirerow"></a>getEntireRow()
Obtient un objet qui représente la ligne entière de la plage (par exemple, si la plage actuelle représente les cellules « B4:E11 », sa valeur `GetEntireRow` est une plage qui représente les lignes « 4:11 »).

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getEntireRow();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "D:F"; 
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeER = range.getEntireRow();
    rangeER.load('address');
    return ctx.sync().then(function() {
        console.log(rangeER.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
Les propriétés de grille de la plage (valeurs, format de nombre, formules) contiennent la valeur `null`, car la plage en question est illimitée.


### <a name="getintersectionanotherrange-range-or-string"></a>getIntersection(anotherRange: Range or string)
Obtient l’objet de plage qui représente l’intersection rectangulaire des plages données.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getIntersection(anotherRange);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|anotherRange|range ou string|Objet de plage ou adresse de plage utilisé pour déterminer l’intersection des plages.|

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getIntersection("D4:G6");
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!D4:F6
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getintersectionornullobjectanotherrange-range-or-string"></a>getIntersectionOrNullObject(anotherRange: range ou string)
Obtient l’objet de plage qui représente l’intersection rectangulaire des plages données. Si aucune intersection n’est trouvée, renvoie un objet Null.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getIntersectionOrNullObject(anotherRange);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|anotherRange|range ou string|Objet de plage ou adresse de plage utilisé pour déterminer l’intersection des plages.|

#### <a name="returns"></a>Retourne
[Range](range.md)

### <a name="getlastcell"></a>getLastCell()
Obtient la dernière cellule de la plage. Par exemple, la dernière cellule de la plage « B2:D5 » est « D5 ».

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getLastCell();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastCell();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getlastcolumn"></a>getLastColumn()
Obtient la dernière colonne de la plage. Par exemple, la dernière colonne de la plage « B2:D5 » est « D2:D5 ».

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getLastColumn();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastColumn();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!F1:F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getlastrow"></a>getLastRow()
Obtient la dernière ligne de la plage. Par exemple, la dernière ligne de la plage « B2:D5 » est « B5:D5 ».

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getLastRow();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastRow();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!A8:F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```



### <a name="getoffsetrangerowoffset-number-columnoffset-number"></a>getOffsetRange(rowOffset: number, columnOffset: number)
Obtient un objet qui représente une plage décalée par rapport à la plage spécifiée. Les dimensions de la plage renvoyée correspondent à cette plage. Si la plage obtenue se retrouve en dehors des limites de grille de la feuille de calcul, une erreur est déclenchée.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getOffsetRange(rowOffset, columnOffset);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|rowOffset|number|Nombre de lignes (positif, négatif ou nul) duquel décaler la plage. Les valeurs positives représentent un décalage vers le bas et les valeurs négatives un décalage vers le haut.|
|columnOffset|number|Nombre de colonnes (positif, négatif ou nul) duquel décaler la plage. Les valeurs positives représentent un décalage vers la droite et les valeurs négatives un décalage vers la gauche.|

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:F6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getOffsetRange(-1,4);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!H3:K5
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getresizedrangedeltarows-number-deltacolumns-number"></a>getResizedRange(deltaRows: nombre, deltaColumns: nombre)
Obtient un objet de plage semblable à l’objet de plage actuel, mais avec le coin inférieur droit développé (ou contracté) selon un certain nombre de lignes et de colonnes.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getResizedRange(deltaRows, deltaColumns);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|deltaRows|number|Nombre de lignes par lequel développer le coin inférieur droit, par rapport à la plage actuelle. Utilisez un nombre positif pour étendre la plage ou un nombre négatif pour la réduire.|
|deltaColumns|number|Nombre de colonnes par lequel développer le coin inférieur droit, par rapport à la plage actuelle. Utilisez un nombre positif pour étendre la plage ou un nombre négatif pour la réduire.|

#### <a name="returns"></a>Retourne
[Range](range.md)

### <a name="getrowrow-number"></a>getRow(row: number)
Obtient une ligne contenue dans la plage.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getRow(row);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|row|number|Numéro de ligne de la plage à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getRow(1);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!A2:F2
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrowsabovecount-number"></a>getRowsAbove(count: nombre)
Obtient un certain nombre de lignes au-dessus de l’objet de plage actuel.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getRowsAbove(count);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|count|number|Facultatif. Nombre de lignes à inclure dans la plage obtenue. En règle générale, utilisez un nombre positif pour créer une plage en dehors de la plage actuelle. Vous pouvez également utiliser un nombre négatif pour créer une plage à l’intérieur de la plage actuelle. La valeur par défaut est 1.|

#### <a name="returns"></a>Renvoie
[Range](range.md)

### <a name="getrowsbelowcount-number"></a>getRowsBelow(count: nombre)
Obtient un certain nombre de lignes en dessous de l’objet de plage actuel.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getRowsBelow(count);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|count|number|Facultatif. Nombre de lignes à inclure dans la plage obtenue. En règle générale, utilisez un nombre positif pour créer une plage en dehors de la plage actuelle. Vous pouvez également utiliser un nombre négatif pour créer une plage à l’intérieur de la plage actuelle. La valeur par défaut est 1.|

#### <a name="returns"></a>Renvoie
[Range](range.md)

### <a name="getusedrangevaluesonly-apisetversion"></a>getUsedRange(valuesOnly: [ApiSet(Version)
Renvoie la plage utilisée d’un objet de plage donné. Si aucune cellule n’est utilisée dans la plage, cette fonction génère une erreur ItemNotFound.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getUsedRange(valuesOnly);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|valuesOnly|[ApiSet(Version|Prend uniquement en compte les cellules avec des valeurs sous forme de cellules utilisées.|

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeUR = range.getUsedRange();
    rangeUR.load('address');
    return ctx.sync().then(function() {
        console.log(rangeUR.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getusedrangeornullobjectvaluesonly-bool"></a>getUsedRangeOrNullObject(valuesOnly: bool)
Renvoie la plage utilisée d’un objet de plage donné. Si aucune cellule n’est utilisée dans la plage, cette fonction renvoie un objet null.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getUsedRangeOrNullObject(valuesOnly);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|valuesOnly|bool|Facultatif. Prend uniquement en compte les cellules avec des valeurs sous forme de cellules utilisées.|

#### <a name="returns"></a>Renvoie
[Range](range.md)

### <a name="getvisibleview"></a>getVisibleView()
Représente les lignes visibles de la plage en cours.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getVisibleView();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[RangeView](rangeview.md)

### <a name="insertshift-string"></a>insert(shift: string)
Insère une cellule ou une plage de cellules dans la feuille de calcul à la place d’une plage donnée et décale les autres cellules pour libérer de l’espace. Renvoie un nouvel objet Range dans l’espace vide qui s’est créé.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.insert(shift);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|Shift|string|Indique la façon dont les cellules doivent être décalées. Les valeurs possibles sont les suivantes : Down (vers le bas), Right (vers la droite)|

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

```js
    
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:F10";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.insert();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="mergeacross-bool"></a>merge(across: bool)
Fusionne la plage de cellules dans une zone de la feuille de calcul.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.merge(across);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|across|bool|Facultatif. Définit la valeur « true » pour fusionner séparément les cellules de chaque ligne de la plage spécifiée. La valeur par défaut est « false ».|

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.merge(true);
    return ctx.sync(); 
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
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.unmerge();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="select"></a>select()
Sélectionne la plage spécifiée dans l’interface utilisateur d’Excel.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.select();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples

```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "F5:F10"; 
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.select();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="unmerge"></a>unmerge()
Annule la fusion de la plage de cellules.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.unmerge();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.unmerge();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

L’exemple ci-dessous utilise l’adresse de la plage pour obtenir l’objet de la plage.

```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8"; 
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

L’exemple ci-dessous utilise une plage nommée pour obtenir l’objet de la plage.

```js

Excel.run(function (ctx) { 
    var rangeName = 'MyRange';
    var range = ctx.workbook.names.getItem(rangeName).range;
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

L’exemple ci-dessous définit le format de nombre, les valeurs et les formules dans une grille 2x3.

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:G7";
    var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
    var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
    var formulas = [[null,null], [null,null], [null,"=G6-G5"]];
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.numberFormat = numberFormat;
    range.values = values;
    range.formulas= formulas;
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
Obtenir la feuille de calcul contenant la plage 

```js
/* This might be broken still - it was broken before because it 
    it was missing 'var', but might still be wrong because of
    getting information without loading properly. */
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var namedItem = names.getItem('MyRange');
    var range = namedItem.range;
    var rangeWorksheet = range.worksheet;
    rangeWorksheet.load('name');
    return ctx.sync().then(function() {
            console.log(rangeWorksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

