# <a name="binding-object-javascript-api-for-excel"></a>Objet Binding (API JavaScript pour Excel)

Représente une liaison Office.js définie dans le classeur.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|id|string|Représente l’identificateur de liaison. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Renvoie le type de la liaison. En lecture seule. Les valeurs possibles sont les suivantes : Range, Table, Text.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Supprime la liaison.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|Renvoie la plage représentée par la liaison. Génère une erreur si la liaison n’est pas du type approprié.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getTable()](#gettable)|[Table](table.md)|Renvoie le tableau représenté par la liaison. Génère une erreur si la liaison n’est pas du type approprié.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getText()](#gettext)|string|Renvoie le texte représenté par la liaison. Génère une erreur si la liaison n’est pas du type approprié.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="delete"></a>delete()
Supprime la liaison.

#### <a name="syntax"></a>Syntaxe
```js
bindingObject.delete();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

### <a name="getrange"></a>getRange()
Renvoie la plage représentée par la liaison. Génère une erreur si la liaison n’est pas du type approprié.

#### <a name="syntax"></a>Syntaxe
```js
bindingObject.getRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
L’exemple ci-dessous utilise l’objet de liaison pour obtenir la plage associée.

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var range = binding.getRange();
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="gettable"></a>getTable()
Renvoie le tableau représenté par la liaison. Génère une erreur si la liaison n’est pas du type approprié.

#### <a name="syntax"></a>Syntaxe
```js
bindingObject.getTable();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Table](table.md)

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var table = binding.getTable();
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


### <a name="gettext"></a>getText()
Renvoie le texte représenté par la liaison. Génère une erreur si la liaison n’est pas du type approprié.

#### <a name="syntax"></a>Syntaxe
```js
bindingObject.getText();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
string

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var text = binding.getText();
    binding.load('text');
    return ctx.sync().then(function() {
        console.log(text);
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
    var binding = ctx.workbook.bindings.getItemAt(0);
    binding.load('type');
    return ctx.sync().then(function() {
        console.log(binding.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
