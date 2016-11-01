# <a name="nameditem-object-(javascript-api-for-excel)"></a>Objet NamedItem (interface API JavaScript pour Excel)

Représente un nom défini pour une plage de cellules ou une valeur. Les noms peuvent être des objets nommés primitifs (comme dans le type ci-dessous), un objet de plage et une référence à une plage. Cet objet peut être utilisé pour obtenir un objet de plage associé aux noms.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|name|string|Nom de l’objet. En lecture seule.|
|type|string|Indique le type de référence associé au nom. En lecture seule. Les valeurs possibles sont les suivantes : String, Integer, Double, Boolean, Range.|
|value|object|Représente la formule à laquelle le nom doit faire référence. Par exemple, =Sheet14!$B$2:$H$12, =4.75, etc. En lecture seule.|
|visible|bool|Indique si l’objet est visible ou non.|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[Range](range.md)|Renvoie l’objet de plage qui est associé au nom. Renvoie une exception si le type de l’élément nommé n’est pas une plage.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## <a name="method-details"></a>Détails des méthodes


### <a name="getrange()"></a>getRange()
Renvoie l’objet de plage qui est associé au nom. Renvoie une exception si le type de l’élément nommé n’est pas une plage.

#### <a name="syntax"></a>Syntaxe
```js
namedItemObject.getRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

Renvoie l’objet de plage qui est associé au nom. Renvoie `null` si le nom n’est pas du type `Range`. Remarque : actuellement, cette API prend uniquement en charge les éléments de classeur inclus dans l’étendue.

```js
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var range = names.getItem('MyRange').getRange();
    range.load('address');
    return ctx.sync().then(function() {
            console.log(range.address);
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
    var names = ctx.workbook.names;
    var namedItem = names.getItem('MyRange');
    namedItem.load('type');
    return ctx.sync().then(function() {
            console.log(namedItem.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```