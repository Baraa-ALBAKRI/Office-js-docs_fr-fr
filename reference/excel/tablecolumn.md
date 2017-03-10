# <a name="tablecolumn-object-javascript-api-for-excel"></a>Objet TableColumn (API JavaScript pour Excel)

Représente une colonne dans un tableau.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|id|int|Renvoie une clé unique qui identifie la colonne du tableau. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|index|int|Renvoie le numéro d’indice de la colonne dans la collection de colonnes du tableau. Avec indice zéro. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Représente le nom de la colonne du tableau.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Une cellule contenant une erreur renvoie la chaîne d’erreur.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|filtre|[Filter](filter.md)|Extrait le filtre appliqué à la colonne. En lecture seule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Supprime la colonne du tableau.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Obtient l’objet de plage associé au corps de données de la colonne.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Obtient l’objet de plage associé à la ligne d’en-tête de la colonne.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|Renvoie l’objet de plage associé à l’intégralité de la colonne.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Obtient l’objet de plage associé à la ligne de total de la colonne.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="delete"></a>delete()
Supprime la colonne du tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableColumnObject.delete();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var column = ctx.workbook.tables.getItem(tableName).columns.getItemAt(2);
    column.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getdatabodyrange"></a>getDataBodyRange()
Obtient l’objet de plage associé au corps de données de la colonne.

#### <a name="syntax"></a>Syntaxe
```js
tableColumnObject.getDataBodyRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var column = ctx.workbook.tables.getItem(tableName).columns.getItemAt(0);
    var dataBodyRange = column.getDataBodyRange();
    dataBodyRange.load('address');
    return ctx.sync().then(function() {
        console.log(dataBodyRange.address);
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getheaderrowrange"></a>getHeaderRowRange()
Obtient l’objet de plage associé à la ligne d’en-tête de la colonne.

#### <a name="syntax"></a>Syntaxe
```js
tableColumnObject.getHeaderRowRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).columns.getItemAt(0);
    var headerRowRange = columns.getHeaderRowRange();
    headerRowRange.load('address');
    return ctx.sync().then(function() {
        console.log(headerRowRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getrange"></a>getRange()
Renvoie l’objet de plage associé à l’intégralité de la colonne.

#### <a name="syntax"></a>Syntaxe
```js
tableColumnObject.getRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).columns.getItemAt(0);
    var columnRange = columns.getRange();
    columnRange.load('address');
    return ctx.sync().then(function() {
        console.log(columnRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="gettotalrowrange"></a>getTotalRowRange()
Obtient l’objet de plage associé à la ligne de total de la colonne.

#### <a name="syntax"></a>Syntaxe
```js
tableColumnObject.getTotalRowRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).columns.getItemAt(0);
    var totalRowRange = columns.getTotalRowRange();
    totalRowRange.load('address');
    return ctx.sync().then(function() {
        console.log(totalRowRange.address);
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
    var column = ctx.workbook.tables.getItem(tableName).columns.getItem(0);
    column.load('index');
    return ctx.sync().then(function() {
        console.log(column.index);
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
    var tableName = 'Table1';
    var tables = ctx.workbook.tables;
    var newValues = [["New"], ["Values"], ["For"], ["New"], ["Column"]];
    var column = ctx.workbook.tables.getItem(tableName).columns.getItemAt(2);
    column.values = newValues;
    column.load('values');
    return ctx.sync().then(function() {
        console.log(column.values);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```