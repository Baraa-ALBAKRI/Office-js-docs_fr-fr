# <a name="tablecolumncollection-object-javascript-api-for-excel"></a>Objet TableColumnCollection (interface API JavaScript pour Excel)

Représente une collection de toutes les colonnes du tableau.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|count|int|Renvoie le nombre de colonnes dans le tableau. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|éléments|[TableColumn[]](tablecolumn.md)|Collection d’objets tableColumn. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[add(index: number, values: (boolean or string or number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableColumn](tablecolumn.md)|Ajoute une nouvelle colonne au tableau.|[1.1, 1.1 nécessite un index plus petit que le nombre total de la colonne ; 1,4 permet que l’index soit facultatif (null ou -1)](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[TableColumn](tablecolumn.md)|Obtient un objet de colonne par son nom ou son ID.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[TableColumn](tablecolumn.md)|Obtient une colonne en fonction de sa position dans la collection.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(key: nombre ou chaîne)](#getitemornullkey-number-or-string)|[TableColumn](tablecolumn.md)|Obtient un objet de colonne par son nom ou son ID. Si la colonne n’existe pas, la propriété isNull de l’objet renvoyé aura la valeur true.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="addindex-number-values-boolean-or-string-or-number"></a>add(index: number, values: (boolean ou string ou number)[][])
Ajoute une nouvelle colonne au tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableColumnCollectionObject.add(index, values);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|index|number|Facultatif. Spécifie la position relative de la nouvelle colonne. Si la valeur est null ou -1, la ligne est ajoutée à la fin. Les colonnes avec un index supérieur seront décalées sur le côté. Avec indice zéro.|
|valeurs|(boolean ou string ou number)[][]|Facultatif. Matrice 2D des valeurs non mises en forme de la colonne du tableau.|

#### <a name="returns"></a>Retourne
[TableColumn](tablecolumn.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    var values = [["Sample"], ["Values"], ["For"], ["New"], ["Column"]];
    var column = tables.getItem("Table1").columns.add(null, values);
    column.load('name');
    return ctx.sync().then(function() {
        console.log(column.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitemkey-number-or-string"></a>getItem(key: number ou string)
Obtient un objet de colonne par son nom ou son ID.

#### <a name="syntax"></a>Syntaxe
```js
tableColumnCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|Key|number ou string| Nom ou ID de la colonne.|

#### <a name="returns"></a>Retourne
[TableColumn](tablecolumn.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var tablecolumn = ctx.workbook.tables.getItem('Table1').columns.getItem(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
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
    var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getitematindex-number"></a>getItemAt(index: number)
Obtient une colonne en fonction de sa position dans la collection.

#### <a name="syntax"></a>Syntaxe
```js
tableColumnCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[TableColumn](tablecolumn.md)

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getitemornullkey-number-or-string"></a>getItemOrNull(key: nombre ou chaîne)
Obtient un objet de colonne par son nom ou son ID. Si la colonne n’existe pas, la propriété isNull de l’objet renvoyé aura la valeur true.

#### <a name="syntax"></a>Syntaxe
```js
tableColumnCollectionObject.getItemOrNull(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|Key|number ou string| Nom ou ID de la colonne.|

#### <a name="returns"></a>Retourne
[TableColumn](tablecolumn.md)

### <a name="loadparam-object"></a>load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### <a name="syntax"></a>Syntaxe
```js
object.load(param);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### <a name="returns"></a>Retourne
void
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

```js
Excel.run(function (ctx) { 
    var tablecolumns = ctx.workbook.tables.getItem('Table1').columns;
    tablecolumns.load('items');
    return ctx.sync().then(function() {
        console.log("tablecolumns Count: " + tablecolumns.count);
        for (var i = 0; i < tablecolumns.items.length; i++)
        {
            console.log(tablecolumns.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```