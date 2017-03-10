# <a name="table-object-javascript-api-for-excel"></a>Objet Table (API JavaScript pour Excel)

Représente un tableau Excel.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|highlightFirstColumn|bool|Indique si la première colonne contient une mise en forme spéciale.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|highlightLastColumn|bool|Indique si la dernière colonne contient une mise en forme spéciale.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|id|int|Renvoie une valeur qui identifie le tableau dans un classeur donné. La valeur de l’identificateur reste identique, même lorsque le tableau est renommé. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Nom du tableau.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showBandedColumns|bool|Indique si les colonnes affichent une mise en forme à bandes dans laquelle la mise en évidence des colonnes impaires diffère de celle des colonnes paires pour faciliter la lecture du tableau.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showBandedRows|bool|Indique si les lignes affichent une mise en forme à bandes dans laquelle la mise en évidence des lignes impaires diffère de celle des lignes paires pour faciliter la lecture du tableau.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showFilterButton|bool|Indique si les boutons de filtre sont visibles dans la partie supérieure de chaque en-tête de colonne. Ce paramètre est autorisé uniquement si le tableau contient une ligne d’en-tête.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showHeaders|bool|Indique si la ligne d’en-tête est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne d’en-tête.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showTotals|bool|Indique si la ligne de total est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne de total.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|style|string|Valeur de constante qui représente le style du tableau. Les valeurs possibles sont les suivantes : TableStyleLight1 à TableStyleLight21, TableStyleMedium1 à TableStyleMedium28, TableStyleStyleDark1 à TableStyleStyleDark11. Vous pouvez également indiquer un style personnalisé présent dans le classeur.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|colonnes|[TableColumnCollection](tablecolumncollection.md)|Représente une collection de toutes les colonnes du tableau. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|Objet Rows|[TableRowCollection](tablerowcollection.md)|Représente une collection de toutes les lignes du tableau. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|tri|[TableSort](tablesort.md)|Représente le tri du tableau. En lecture seule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|feuille de calcul|[Worksheet](worksheet.md)|Feuille de calcul contenant le tableau actif. En lecture seule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[clearFilters()](#clearfilters)|void|Supprime tous les filtres appliqués actuellement sur le tableau.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[convertToRange()](#converttorange)|[Range](range.md)|Convertit le tableau en plage normale de cellules. Toutes les données sont conservées.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[delete()](#delete)|void|Supprime le tableau.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Obtient l’objet de plage associé au corps de données du tableau.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Obtient l’objet de plage associé à la ligne d’en-tête du tableau.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|Renvoie l’objet de plage associé à l’intégralité du tableau.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Renvoie l’objet de plage associé à la ligne de total du tableau.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[reapplyFilters()](#reapplyfilters)|void|Applique de nouveau tous les filtres actuellement appliqués sur le tableau.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="clearfilters"></a>clearFilters()
Supprime tous les filtres appliqués actuellement sur le tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableObject.clearFilters();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

### <a name="converttorange"></a>convertToRange()
Convertit le tableau en plage normale de cellules. Toutes les données sont conservées.

#### <a name="syntax"></a>Syntaxe
```js
tableObject.convertToRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.convertToRange();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="delete"></a>delete()
Supprime le tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableObject.delete();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getdatabodyrange"></a>getDataBodyRange()
Obtient l’objet de plage associé au corps de données du tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableObject.getDataBodyRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableDataRange = table.getDataBodyRange();
    tableDataRange.load('address')
    return ctx.sync().then(function() {
            console.log(tableDataRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getheaderrowrange"></a>getHeaderRowRange()
Obtient l’objet de plage associé à la ligne d’en-tête du tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableObject.getHeaderRowRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableHeaderRange = table.getHeaderRowRange();
    tableHeaderRange.load('address');
    return ctx.sync().then(function() {
        console.log(tableHeaderRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrange"></a>getRange()
Renvoie l’objet de plage associé à l’intégralité du tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableObject.getRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableRange = table.getRange();
    tableRange.load('address');    
    return ctx.sync().then(function() {
            console.log(tableRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="gettotalrowrange"></a>getTotalRowRange()
Renvoie l’objet de plage associé à la ligne de total du tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableObject.getTotalRowRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableTotalsRange = table.getTotalRowRange();
    tableTotalsRange.load('address');    
    return ctx.sync().then(function() {
            console.log(tableTotalsRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="reapplyfilters"></a>reapplyFilters()
Applique de nouveau tous les filtres actuellement appliqués sur le tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableObject.reapplyFilters();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

Obtenir un tableau par son nom 

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.load('index')
    return ctx.sync().then(function() {
            console.log(table.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Obtenir un tableau par son indice

```js
Excel.run(function (ctx) { 
    var index = 0;
    var table = ctx.workbook.tables.getItemAt(0);
    table.load('id')
    return ctx.sync().then(function() {
            console.log(table.id);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Définir le style du tableau 

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.name = 'Table1-Renamed';
    table.showTotals = false;
    table.style = 'TableStyleMedium2';
    table.load('tableStyle');
    return ctx.sync().then(function() {
            console.log(table.style);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
