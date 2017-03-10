# <a name="nameditem-object-javascript-api-for-excel"></a>Objet NamedItem (API JavaScript pour Excel)

Représente un nom défini pour une plage de cellules ou une valeur. Les noms peuvent être des objets nommés primitifs (comme dans le type ci-dessous), un objet de plage ou une référence à une plage. Cet objet peut être utilisé pour obtenir l’objet de plage associé à des noms.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|comment|string|Représente le commentaire associé à ce nom.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Nom de l’objet. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|scope|string|Indique si le nom est étendu au classeur ou à une feuille de calcul spécifique. En lecture seule. Les valeurs possibles sont les suivantes : Equal, Greater, GreaterEqual, Less, LessEqual, NotEqual.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Indique le type de la valeur renvoyée par la formule du nom. En lecture seule. Les valeurs possibles sont les suivantes : String, Integer, Double, Boolean, Range.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|value|object|Représente la valeur calculée par la formule du nom. Pour une plage nommée, renvoie l’adresse de la plage. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|Indique si l’objet est visible ou non.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|feuille de calcul|[Worksheet](worksheet.md)|Renvoie la feuille de calcul dans laquelle est inclus l’élément nommé. Génère une erreur si les éléments sont inclus dans le classeur à la place. En lecture seule.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|worksheetOrNullObject|[Worksheet](worksheet.md)|Renvoie la feuille de calcul dans laquelle est inclus l’élément nommé. Renvoie un objet null si l’élément est inclus dans le classeur à la place. En lecture seule.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Supprime le nom donné.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|Renvoie l’objet de plage qui est associé au nom. Renvoie une erreur si le type de l’élément nommé n’est pas une plage.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRangeOrNullObject()](#getrangeornullobject)|[Range](range.md)|Renvoie l’objet de plage qui est associé au nom. Renvoie un objet null si le type de l’élément nommé n’est pas une plage.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="delete"></a>delete()
Supprime le nom donné.

#### <a name="syntax"></a>Syntaxe
```js
namedItemObject.delete();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

### <a name="getrange"></a>getRange()
Renvoie l’objet de plage qui est associé au nom. Renvoie une erreur si le type de l’élément nommé n’est pas une plage.

#### <a name="syntax"></a>Syntaxe
```js
namedItemObject.getRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

Renvoie l’objet de plage qui est associé au nom. Renvoie `null` si le nom n’est pas du type `Range`. Remarque : actuellement, cette API prend uniquement en charge les éléments de classeur inclus dans l’étendue.**

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


### <a name="getrangeornullobject"></a>getRangeOrNullObject()
Renvoie l’objet de plage qui est associé au nom. Renvoie un objet null si le type de l’élément nommé n’est pas une plage.

#### <a name="syntax"></a>Syntaxe
```js
namedItemObject.getRangeOrNullObject();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)
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
