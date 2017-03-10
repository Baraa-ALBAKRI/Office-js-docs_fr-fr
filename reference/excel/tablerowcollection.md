# <a name="tablerowcollection-object-javascript-api-for-excel"></a>Objet TableRowCollection (API JavaScript pour Excel)

Représente une collection de toutes les lignes du tableau.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|count|int|Renvoie le nombre de lignes du tableau. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|éléments|[TableRow[]](tablerow.md)|Collection d’objets tableRow. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[add(index: number, values: (boolean or string or number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableRow](tablerow.md)|Ajoute une ou plusieurs lignes dans le tableau. L’objet renvoyé sera placé en premier dans les lignes récemment ajoutées.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Obtient le nombre de lignes dans le tableau.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[TableRow](tablerow.md)|Obtient une ligne en fonction de sa position dans la collection.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="addindex-number-values-boolean-or-string-or-number"></a>add(index: number, values: (boolean ou string ou number)[][])
Ajoute une ou plusieurs lignes dans le tableau. L’objet renvoyé sera placé en premier dans les lignes récemment ajoutées.

#### <a name="syntax"></a>Syntaxe
```js
tableRowCollectionObject.add(index, values);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|index|number|Facultatif. Spécifie la position relative de la nouvelle ligne. Si la valeur est null ou -1, la ligne est ajoutée à la fin. Toutes les lignes en dessous de la ligne insérée sont déplacées vers le bas. Avec indice zéro.|
|values|(boolean ou string ou number)[][]|Facultatif. Matrice 2D des valeurs non mises en forme de la ligne du tableau.|

#### <a name="returns"></a>Retourne
[TableRow](tablerow.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    var values = [["Sample", "Values", "For", "New", "Row"]];
    var row = tables.getItem("Table1").rows.add(null, values);
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

### <a name="getcount"></a>getCount()
Obtient le nombre de lignes dans le tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableRowCollectionObject.getCount();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Renvoie
int

### <a name="getitematindex-number"></a>getItemAt(index: number)
Obtient une ligne en fonction de sa position dans la collection.

#### <a name="syntax"></a>Syntaxe
```js
tableRowCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[TableRow](tablerow.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var tablerow = ctx.workbook.tables.getItem('Table1').rows.getItemAt(0);
    tablerow.load('name');
    return ctx.sync().then(function() {
            console.log(tablerow.name);
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
    var tablerows = ctx.workbook.tables.getItem('Table1').rows;
    tablerows.load('items');
    return ctx.sync().then(function() {
        console.log("tablerows Count: " + tablerows.count);
        for (var i = 0; i < tablerows.items.length; i++)
        {
            console.log(tablerows.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```