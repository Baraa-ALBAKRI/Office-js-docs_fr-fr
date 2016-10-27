# <a name="tablecollection-object-(javascript-api-for-excel)"></a>Objet TableCollection (interface API JavaScript pour Excel)

Représente une collection de tous les tableaux du classeur.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|count|int|Renvoie le nombre de tableaux dans le classeur. En lecture seule.|
|items|[Table[]](table.md)|Collection d’objets de tableau. En lecture seule.|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[add(address: string, hasHeaders: bool)](#addaddress-string-hasheaders-bool)|[Table](table.md)|Crée un tableau. L’adresse de la source de la plage détermine la feuille de calcul dans laquelle le tableau sera ajouté. Si l’ajout ne peut être effectué (par exemple, parce que l’adresse n’est pas valide, ou parce que le tableau empiéterait sur un autre tableau), un message d’erreur apparaît.|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[Table](table.md)|Obtient un tableau à l’aide de son nom ou de son ID.|
|[getItemAt(index: number)](#getitematindex-number)|[Table](table.md)|Obtient un tableau en fonction de sa position dans la collection.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## <a name="method-details"></a>Détails des méthodes


### <a name="add(address:-string,-hasheaders:-bool)"></a>add(address: string, hasHeaders: bool)
Crée un tableau. L’adresse de la source de la plage détermine la feuille de calcul dans laquelle le tableau sera ajouté. Si l’ajout ne peut être effectué (par exemple, parce que l’adresse n’est pas valide, ou parce que le tableau empiéterait sur un autre tableau), un message d’erreur apparaît.

#### <a name="syntax"></a>Syntaxe
```js
tableCollectionObject.add(address, hasHeaders);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|address|string|Adresse ou nom de l’objet de plage représentant la source de données. Notez que l’adresse de la plage doit inclure la feuille de calcul dans laquelle le tableau doit être ajouté. Exemple `Sheet1!A1:D4`.|
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

### <a name="getitem(key:-number-or-string)"></a>getItem(key: number or string)
Obtient un tableau à l’aide de son nom ou de son ID.

#### <a name="syntax"></a>Syntaxe
```js
tableCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|Key|number or string|Nom ou ID du tableau à récupérer.|

#### <a name="returns"></a>Retourne
[Table](table.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
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


#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItemAt(0);
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


### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
Obtient un tableau en fonction de sa position dans la collection.

#### <a name="syntax"></a>Syntaxe
```js
tableCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[Table](table.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItemAt(0);
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
    var tables = ctx.workbook.tables;
    tables.load('items');
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
