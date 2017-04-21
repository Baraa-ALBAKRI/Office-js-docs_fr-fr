# <a name="worksheetcollection-object-javascript-api-for-excel"></a>Objet WorksheetCollection (API JavaScript pour Excel)

Représente une collection d’objets de feuille de calcul qui font partie du classeur.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|éléments|[Worksheet[]](worksheet.md)|Collection d’objets de feuille de calcul. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[add(name: string)](#addname-string)|[Worksheet](worksheet.md)|Ajoute une nouvelle feuille de calcul au classeur. La feuille de calcul est ajoutée à la fin des feuilles de calcul existantes. Si vous souhaitez activer la feuille de calcul nouvellement ajoutée, appelez la méthode .activate() pour cette feuille.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getActiveWorksheet()](#getactiveworksheet)|[Worksheet](worksheet.md)|Obtient la feuille de calcul active du classeur.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount(visibleOnly: bool)](#getcountvisibleonly-bool)|int|Obtient le nombre de feuilles de calcul dans la collection.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[Worksheet](worksheet.md)|Obtient un objet de feuille de calcul à l’aide de son nom ou de son ID.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: string)](#getitemornullobjectkey-string)|[Feuille de calcul](worksheet.md)|Obtient un objet de feuille de calcul à l’aide de son nom ou de son ID. Si la feuille de calcul n’existe pas, renvoie un objet null.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="addname-string"></a>add(name: string)
Ajoute une nouvelle feuille de calcul au classeur. La feuille de calcul est ajoutée à la fin des feuilles de calcul existantes. Si vous souhaitez activer la feuille de calcul nouvellement ajoutée, appelez la méthode .activate() pour cette feuille.

#### <a name="syntax"></a>Syntaxe
```js
worksheetCollectionObject.add(name);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|name|string|Facultatif. Nom de la feuille de calcul à ajouter. Si cette propriété est définie, le nom doit être unique. Si cette propriété n’est pas définie, Excel détermine le nom de la nouvelle feuille de calcul.|

#### <a name="returns"></a>Retourne
[Worksheet](worksheet.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sample Name';
    var worksheet = ctx.workbook.worksheets.add(wSheetName);
    worksheet.load('name');
    return ctx.sync().then(function() {
        console.log(worksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getactiveworksheet"></a>getActiveWorksheet()
Obtient la feuille de calcul active du classeur.

#### <a name="syntax"></a>Syntaxe
```js
worksheetCollectionObject.getActiveWorksheet();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Worksheet](worksheet.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) {  
    var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
    activeWorksheet.load('name');
    return ctx.sync().then(function() {
            console.log(activeWorksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcountvisibleonly-bool"></a>getCount(visibleOnly: bool)
Obtient le nombre de feuilles de calcul dans la collection.

#### <a name="syntax"></a>Syntaxe
```js
worksheetCollectionObject.getCount(visibleOnly);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|visibleOnly|bool|Facultatif. Renvoie des feuilles de calcul visibles uniquement si la valeur est définie sur true. |

#### <a name="returns"></a>Retourne
int

### <a name="getitemkey-string"></a>getItem(key: string)
Obtient un objet de feuille de calcul à l’aide de son nom ou de son ID.

#### <a name="syntax"></a>Syntaxe
```js
worksheetCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|Key|string|Nom ou ID de la feuille de calcul.|

#### <a name="returns"></a>Retourne
[Worksheet](worksheet.md)

### <a name="getitemornullobjectkey-string"></a>getItemOrNullObject(key: string)
Obtient un objet de feuille de calcul à l’aide de son nom ou de son ID. Si la feuille de calcul n’existe pas, renvoie un objet null.

#### <a name="syntax"></a>Syntaxe
```js
worksheetCollectionObject.getItemOrNullObject(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|Key|string|Nom ou ID de la feuille de calcul.|

#### <a name="returns"></a>Retourne
[Worksheet](worksheet.md)
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés
```js
Excel.run(function (ctx) { 
    var worksheets = ctx.workbook.worksheets;
    worksheets.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < worksheets.items.length; i++)
        {
            console.log(worksheets.items[i].name);
            console.log(worksheets.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
