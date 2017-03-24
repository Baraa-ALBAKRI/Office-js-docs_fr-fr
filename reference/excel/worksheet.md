# <a name="worksheet-object-javascript-api-for-excel"></a>Objet Worksheet (API JavaScript pour Excel)

Une feuille de calcul Excel est une grille de cellules. Elle peut contenir des données, des tableaux, des graphiques, etc.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|id|string|Renvoie une valeur qui permet d’identifier la feuille de calcul de façon unique dans un classeur donné. La valeur de l’identificateur reste identique, même lorsque la feuille de calcul est renommée ou déplacée. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Nom complet de la feuille de calcul.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|Position|int|Position de la feuille de calcul au sein du classeur (sur une base zéro).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visibility|string|Visibilité de la feuille de calcul. Les valeurs possibles sont les suivantes : Visible (visible), Hidden (masquée), VeryHidden (très masquée).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|charts|[ChartCollection](chartcollection.md)|Renvoie une collection de graphiques qui font partie de la feuille de calcul. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|Noms|[NamedItemCollection](nameditemcollection.md)|Collection de noms inclus dans l’étendue de la feuille de calcul active. En lecture seule.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|pivotTables|[PivotTableCollection](pivottablecollection.md)|Collection de tableaux croisés dynamiques qui font partie de la feuille de calcul. En lecture seule.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|protection|[WorksheetProtection](worksheetprotection.md)|Renvoie un objet de protection de feuille pour une feuille de calcul. En lecture seule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|tables|[TableCollection](tablecollection.md)|Collection de tableaux qui font partie de la feuille de calcul. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[activate()](#activate)|void|Active la feuille de calcul dans l’interface utilisateur Excel.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[delete()](#delete)|void|Supprime la feuille de calcul du classeur.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Renvoie l’objet de plage qui contient une cellule donnée sur la base des numéros de ligne et de colonne. La cellule peut se trouver en dehors des limites de ses plages parent, pour peu qu’elle reste dans la grille de la feuille de calcul.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange(address: string)](#getrangeaddress-string)|[Range](range.md)|Renvoie l’objet de plage spécifié par son nom ou son adresse.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRange(valuesOnly: bool)](#getusedrangevaluesonly-bool)|[Range](range.md)|La plage utilisée est la plus petite plage qui englobe toutes les cellules auxquelles une valeur ou un format est affecté. Si la feuille de calcul entière est vide, cette fonction renvoie la cellule supérieure gauche (c'est-à-dire qu’elle ne génère *pas* d’erreur).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRangeOrNullObject(valuesOnly: bool)](#getusedrangeornullobjectvaluesonly-bool)|[Range](range.md)|La plage utilisée est la plus petite plage qui englobe toutes les cellules auxquelles une valeur ou un format est affecté. Si la feuille de calcul entière est vide, cette fonction renvoie un objet null.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="activate"></a>activate()
Active la feuille de calcul dans l’interface utilisateur Excel.

#### <a name="syntax"></a>Syntaxe
```js
worksheetObject.activate();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.activate();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="delete"></a>delete()
Supprime la feuille de calcul du classeur.

#### <a name="syntax"></a>Syntaxe
```js
worksheetObject.delete();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcellrow-number-column-number"></a>getCell(row: number, column: number)
Renvoie l’objet de plage qui contient une cellule donnée sur la base des numéros de ligne et de colonne. La cellule peut se trouver en dehors des limites de ses plages parent, pour peu qu’elle reste dans la grille de la feuille de calcul.

#### <a name="syntax"></a>Syntaxe
```js
worksheetObject.getCell(row, column);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|row|number|Numéro de ligne de la cellule à récupérer. Avec indice zéro.|
|column|number|Numéro de colonne de la cellule à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Renvoie
[Range](range.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var cell = worksheet.getCell(0,0);
    cell.load('address');
    return ctx.sync().then(function() {
        console.log(cell.address);
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrangeaddress-string"></a>getRange(address: string)
Renvoie l’objet de plage spécifié par son nom ou son adresse.

#### <a name="syntax"></a>Syntaxe
```js
worksheetObject.getRange(address);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|address|string|Facultatif. Adresse ou nom de la plage. Si cette propriété n’est pas définie, la plage de la feuille de calcul toute entière est renvoyée.|

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
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
    var sheetName = "Sheet1";
    var rangeName = 'MyRange';
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeName);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getusedrangevaluesonly-bool"></a>getUsedRange(valuesOnly: bool)
La plage utilisée est la plus petite plage qui englobe toutes les cellules auxquelles une valeur ou un format est affecté. Si la feuille de calcul entière est vide, cette fonction renvoie la cellule supérieure gauche (c'est-à-dire qu’elle ne génère *pas* d’erreur).

#### <a name="syntax"></a>Syntaxe
```js
worksheetObject.getUsedRange(valuesOnly);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|valuesOnly|[ApiSet(Version|Prend uniquement en compte les cellules avec des valeurs sous forme de cellules utilisées (ignore la mise en forme).|

#### <a name="returns"></a>Renvoie
[Range](range.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    var usedRange = worksheet.getUsedRange();
    usedRange.load('address');
    return ctx.sync().then(function() {
            console.log(usedRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getusedrangeornullobjectvaluesonly-bool"></a>getUsedRangeOrNullObject(valuesOnly: bool)
La plage utilisée est la plus petite plage qui englobe toutes les cellules auxquelles une valeur ou un format est affecté. Si la feuille de calcul entière est vide, cette fonction renvoie un objet null.

#### <a name="syntax"></a>Syntaxe
```js
worksheetObject.getUsedRangeOrNullObject(valuesOnly);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|valuesOnly|bool|Facultatif. Prend uniquement en compte les cellules avec des valeurs sous forme de cellules utilisées.|

#### <a name="returns"></a>Renvoie
[Range](range.md)
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

Obtenir les propriétés de la feuille de calcul à partir du nom de la feuille

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.load('position')
    return ctx.sync().then(function() {
            console.log(worksheet.position);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Définir la position de la feuille de calcul 

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.position = 2;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
