# <a name="binding-object-(javascript-api-for-excel)"></a>Objet Binding (interface API JavaScript pour Excel)

Représente une liaison Office.js définie dans le classeur.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|id|string|Représente l’identificateur de liaison. En lecture seule.|
|type|string|Renvoie le type de la liaison. En lecture seule. Les valeurs possibles sont les suivantes : Range, Table, Text.|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[Range](range.md)|Renvoie la plage représentée par la liaison. Génère une erreur si la liaison n’est pas du type approprié.|
|[getTable()](#gettable)|[Table](table.md)|Renvoie la table représentée par la liaison. Génère une erreur si la liaison n’est pas du type approprié.|
|[getText()](#gettext)|chaîne|Renvoie le texte représenté par la liaison. Génère une erreur si la liaison n’est pas du type approprié.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## <a name="method-details"></a>Détails des méthodes


### <a name="getrange()"></a>getRange()
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
L’exemple ci-dessous utilise un objet de liaison pour obtenir la plage associée.

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


### <a name="gettable()"></a>getTable()
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


### <a name="gettext()"></a>getText()
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
    ctx.load('text');
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


### <a name="load(param:-object)"></a>load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### <a name="syntax"></a>Syntaxe
```js
object.load(param);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, accepte un objet [loadOption](loadoption.md).|

#### <a name="returns"></a>Renvoie
void
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
