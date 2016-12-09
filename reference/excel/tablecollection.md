# <a name="tablecollection-object-javascript-api-for-excel"></a>Objet TableCollection (interface API JavaScript pour Excel)

Représente une collection de tous les tableaux du classeur.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|count|int|Renvoie le nombre de tableaux dans le classeur. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|éléments|[Table[]](table.md)|Collection d’objets de tableau. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[add(address: plage ou chaîne, hasHeaders: valeur booléenne)](#addaddress-range-or-string-hasheaders-bool)|[Table](table.md)|Crée un tableau L’adresse de la source ou de l’objet de la plage détermine la feuille de calcul dans laquelle le tableau sera ajouté. Si l’ajout ne peut être effectué (par exemple, parce que l’adresse n’est pas valide, ou parce que le tableau empiéterait sur un autre tableau), un message d’erreur apparaît.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[Table](table.md)|Obtient un tableau à l’aide de son nom ou de son ID.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Table](table.md)|Obtient un tableau en fonction de sa position dans la collection.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(key: nombre ou chaîne)](#getitemornullkey-number-or-string)|[Table](table.md)|Obtient un tableau à l’aide de son nom ou de son ID. Si le tableau n’existe pas, la propriété isNull de l’objet renvoyé aura la valeur true.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="addaddress-range-or-string-hasheaders-bool"></a>add(address: plage ou chaîne, hasHeaders: valeur booléenne)
Crée un tableau L’adresse de la source ou de l’objet de la plage détermine la feuille de calcul dans laquelle le tableau sera ajouté. Si l’ajout ne peut être effectué (par exemple, parce que l’adresse n’est pas valide, ou parce que le tableau empiéterait sur un autre tableau), un message d’erreur apparaît.

#### <a name="syntax"></a>Syntaxe
```js
tableCollectionObject.add(address, hasHeaders);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|adresse|range ou string|Objet de plage ou nom/adresse (chaîne) de la plage représentant la source de données. Si l’adresse ne contient pas de nom de feuille, la feuille ouverte est utilisée. Ensemble de conditions requises 1.1 pour le paramètre de chaîne ; 1.3 pour accepter un objet de plage.|
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

### <a name="getitemkey-number-or-string"></a>getItem(key: number ou string)
Obtient un tableau à l’aide de son nom ou de son ID.

#### <a name="syntax"></a>Syntaxe
```js
tableCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
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
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[Table](table.md)

#### <a name="examples"></a>範例

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


### <a name="getitemornullkey-number-or-string"></a>getItemOrNull(key: nombre ou chaîne)
Obtient un tableau à l’aide de son nom ou de son ID. Si le tableau n’existe pas, la propriété isNull de l’objet renvoyé aura la valeur true.

#### <a name="syntax"></a>Syntaxe
```js
tableCollectionObject.getItemOrNull(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|Key|number or string|Nom ou ID du tableau à récupérer.|

#### <a name="returns"></a>Retourne
[Table](table.md)

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