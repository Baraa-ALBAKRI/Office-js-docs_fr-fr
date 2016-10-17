# <a name="table-object-(javascript-api-for-excel)"></a>Objet Table (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Excel pour iOS, Office 2016_

Représente un tableau Excel.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|id|int|Renvoie une valeur qui identifie le tableau dans un classeur donné. La valeur de l’identificateur reste identique, même lorsque le tableau est renommé. En lecture seule.|
|name|string|Nom du tableau.|
|showHeaders|bool|Indique si la ligne d’en-tête est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne d’en-tête.|
|showTotals|bool|Indique si la ligne de total est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne de total.|
|style|string|Valeur de constante qui représente le style du tableau. Les valeurs possibles sont les suivantes : TableStyleLight1 à TableStyleLight21, TableStyleMedium1 à TableStyleMedium28, TableStyleStyleDark1 à TableStyleStyleDark11. Vous pouvez également indiquer un style personnalisé présent dans le classeur.|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type   |Description|
|:---------------|:--------|:----------|
|colonnes|[TableColumnCollection](tablecolumncollection.md)|Représente une collection de toutes les colonnes du tableau. En lecture seule.|
|lignes|[TableRowCollection](tablerowcollection.md)|Représente une collection de toutes les lignes du tableau. En lecture seule.|
|tri|[TableSort](tablesort.md)|Représente la configuration de tri du tableau. En lecture seule.|
|feuille de calcul|[Worksheet](worksheet.md)|Feuille de calcul contenant le tableau actuel. En lecture seule.|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[clearFilters()](#clearfilters)|void|Supprime tous les filtres appliqués actuellement sur le tableau.|
|[convertToRange()](#converttorange)|[Range](range.md)|Convertit le tableau en plage normale de cellules. Toutes les données sont conservées.|
|[delete()](#delete)|void|Supprime le tableau.|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Obtient l’objet de plage associé au corps de données du tableau.|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Obtient l’objet de plage associé à la ligne d’en-tête du tableau.|
|[getRange()](#getrange)|[Range](range.md)|Renvoie l’objet de plage associé à l’intégralité du tableau.|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Obtient l’objet de plage associé à la ligne de total du tableau.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|
|[reapplyFilters()](#reapplyfilters)|void|Applique de nouveau tous les filtres actuellement appliqués sur le tableau.|

## <a name="method-details"></a>Détails des méthodes


### <a name="clearfilters()"></a>clearFilters()
Supprime tous les filtres appliqués actuellement sur le tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableObject.clearFilters();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

### <a name="converttorange()"></a>convertToRange()
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

### <a name="delete()"></a>delete()
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


### <a name="getdatabodyrange()"></a>getDataBodyRange()
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

### <a name="getheaderrowrange()"></a>getHeaderRowRange()
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


### <a name="getrange()"></a>getRange()
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


### <a name="gettotalrowrange()"></a>getTotalRowRange()
Obtient l’objet de plage associé à la ligne de total du tableau.

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
    table.name('name')
    return ctx.sync().then(function() {
            console.log(table.name);
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
    table.tableStyle = 'TableStyleMedium2';
    table.load('tableStyle');
    return ctx.sync().then(function() {
            console.log(table.tableStyle);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
