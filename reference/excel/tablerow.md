# <a name="tablerow-object-javascript-api-for-excel"></a>Objet TableRow (API JavaScript pour Excel)

Représente une ligne d’un tableau.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|index|int|Renvoie le numéro d’indice de la ligne dans la collection de lignes du tableau. Avec indice zéro. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Une cellule contenant une erreur renvoie la chaîne d’erreur.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Supprime la ligne du tableau.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|Renvoie l’objet de plage associé à la ligne entière.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="delete"></a>delete()
Supprime la ligne du tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableRowObject.delete();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var row = ctx.workbook.tables.getItem(tableName).rows.getItemAt(2);
    row.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrange"></a>getRange()
Renvoie l’objet de plage associé à la ligne entière.

#### <a name="syntax"></a>Syntaxe
```js
tableRowObject.getRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var row = ctx.workbook.tables.getItem(tableName).rows.getItemAt(0);
    var rowRange = row.getRange();
    rowRange.load('address');
    return ctx.sync().then(function() {
        console.log(rowRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var row = ctx.workbook.tables.getItem(tableName).rows.getItem(0);
    row.load('index');
    return ctx.sync().then(function() {
        console.log(row.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    var newValues = [["New", "Values", "For", "New", "Row"]];
    var tableName = 'Table1';
    var row = ctx.workbook.tables.getItem(tableName).rows.getItemAt(2);
    row.values = newValues;
    row.load('values');
    return ctx.sync().then(function() {
        console.log(row.values);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```