# <a name="tablecolumncollection-object-(javascript-api-for-excel)"></a>Objet TableColumnCollection (interface API JavaScript pour Excel)

Représente une collection de toutes les colonnes du tableau.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|count|int|Renvoie le nombre de colonnes du tableau. En lecture seule.|
|items|[TableColumn[]](tablecolumn.md)|Collection d’objets tableColumn. En lecture seule.|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[add(index: number, values: (boolean or string or number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableColumn](tablecolumn.md)|Ajoute une nouvelle colonne au tableau.|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[TableColumn](tablecolumn.md)|Obtient un objet de colonne par son nom ou son ID.|
|[getItemAt(index: number)](#getitematindex-number)|[TableColumn](tablecolumn.md)|Obtient une colonne en fonction de sa position dans la collection.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## <a name="method-details"></a>Détails des méthodes


### <a name="add(index:-number,-values:-(boolean-or-string-or-number)[][])"></a>add(index: number, values: (boolean ou string ou number)[][])
Ajoute une nouvelle colonne au tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableColumnCollectionObject.add(index, values);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number|Spécifie la position relative de la nouvelle colonne. La colonne qui se trouvait précédemment à cette position est décalée vers la droite. La valeur d’indice doit être égale ou inférieure à celle de la dernière colonne, afin qu’elle n’ajoute pas de colonne à la fin du tableau. Avec indice zéro.|
|values|(boolean ou string ou number)[][]|Facultatif. Matrice 2D des valeurs non mises en forme de la colonne du tableau.|

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


### <a name="getitem(key:-number-or-string)"></a>getItem(key: number ou string)
Obtient un objet de colonne par son nom ou son ID.

#### <a name="syntax"></a>Syntaxe
```js
tableColumnCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|Key|number ou string| Nom ou ID de la colonne.|

#### <a name="returns"></a>Retourne
[TableColumn](tablecolumn.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItem(0);
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

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
Obtient une colonne en fonction de sa position dans la collection.

#### <a name="syntax"></a>Syntaxe
```js
tableColumnCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[TableColumn](tablecolumn.md)

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
    var tablecolumns = ctx.workbook.tables.getItem['Table1'].columns;
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
