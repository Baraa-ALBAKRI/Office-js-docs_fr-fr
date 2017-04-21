# <a name="tablecollection-object-javascript-api-for-excel"></a>Objet TableCollection (API JavaScript pour Excel)

Représente une collection de tous les tableaux qui font partie du classeur ou de la feuille de calcul, en fonction de la méthode d’appel.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|count|int|Renvoie le nombre de tableaux dans le classeur. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|éléments|[Table[]](table.md)|Collection d’objets de tableau. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[add(address: [object, hasHeaders: bool)](#addaddress-object-hasheaders-bool)|[Table](table.md)|Crée un tableau L’adresse de la source ou de l’objet de la plage détermine la feuille de calcul dans laquelle le tableau sera ajouté. Si l’ajout ne peut être effectué (par exemple, parce que l’adresse n’est pas valide, ou parce que le tableau empiéterait sur un autre tableau), un message d’erreur apparaît.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Obtient le nombre de tableaux de la collection.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: number ou string)](#getitemkey-number-or-string)|[Table](table.md)|Obtient un tableau à l’aide de son nom ou de son ID.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Table](table.md)|Obtient un tableau en fonction de sa position dans la collection.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: number ou string)](#getitemornullobjectkey-number-or-string)|[Tableau](table.md)|Obtient un tableau à l’aide de son nom ou de son ID. Si le tableau n’existe pas, renvoie un objet null.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="addaddress-object-hasheaders-bool"></a>add(address: [object, hasHeaders: bool)
Crée un tableau L’adresse de la source ou de l’objet de la plage détermine la feuille de calcul dans laquelle le tableau sera ajouté. Si l’ajout ne peut être effectué (par exemple, parce que l’adresse n’est pas valide, ou parce que le tableau empiéterait sur un autre tableau), un message d’erreur apparaît.

#### <a name="syntax"></a>Syntaxe
```js
tableCollectionObject.add(address, hasHeaders);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|address|[object|Objet de plage ou nom/adresse (string) de la plage représentant la source de données. Si l’adresse ne contient pas de nom de feuille, la feuille ouverte est utilisée. Pour 1.1, utilisez le paramètre string ; pour 1.3 vous pouvez également utiliser l’objet Range.|
|hasHeaders|bool|Valeur booléenne qui indique si les données importées disposent d’étiquettes de colonne. Si la source ne contient pas d’en-têtes (autrement dit, lorsque cette propriété est définie sur false), Excel génère automatiquement un en-tête et décale les données d’une ligne vers le bas.|

#### <a name="returns"></a>Retourne
[Table](table.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.add('Sheet1!A1:E7', true);
    table.load('name');
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

### <a name="getcount"></a>getCount()
Obtient le nombre de tableaux de la collection.

#### <a name="syntax"></a>Syntaxe
```js
tableCollectionObject.getCount();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Renvoie
int

### <a name="getitemkey-number-or-string"></a>getItem(key: number ou string)
Obtient un tableau à l’aide de son nom ou de son ID.

#### <a name="syntax"></a>Syntaxe
```js
tableCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|Key|number or string|Nom ou ID du tableau à récupérer.|

#### <a name="returns"></a>Retourne
[Table](table.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.load('name');
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


#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItemAt(0);
    table.load('name');
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


### <a name="getitematindex-number"></a>getItemAt(index: number)
Obtient un tableau en fonction de sa position dans la collection.

#### <a name="syntax"></a>Syntaxe
```js
tableCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[Table](table.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItemAt(0);
    table.load('name');
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


### <a name="getitemornullobjectkey-number-or-string"></a>getItemOrNullObject(key: number ou string)
Obtient un tableau à l’aide de son nom ou de son ID. Si le tableau n’existe pas, renvoie un objet null.

#### <a name="syntax"></a>Syntaxe
```js
tableCollectionObject.getItemOrNullObject(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|Key|number or string|Nom ou ID du tableau à récupérer.|

#### <a name="returns"></a>Retourne
[Table](table.md)
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    tables.load();
    return ctx.sync().then(function() {
        console.log("tables Count: " + tables.count);
        for (var i = 0; i < tables.items.length; i++)
        {
            console.log(tables.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Obtenir le nombre de tableaux

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    tables.load('count');
    return ctx.sync().then(function() {
        console.log(tables.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```