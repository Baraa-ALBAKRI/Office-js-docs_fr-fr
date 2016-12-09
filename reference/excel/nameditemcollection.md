# <a name="nameditemcollection-object-javascript-api-for-excel"></a>Objet NamedItemCollection (interface API JavaScript pour Excel)

Collection de tous les objets NamedItem du classeur.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|éléments|[NamedItem[]](nameditem.md)|Collection d’objets NamedItem. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|Obtient un objet NamedItem à l’aide de son nom.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(name: chaîne)](#getitemornullname-string)|[NamedItem](nameditem.md)|Obtient un objet NamedItem à l’aide de son nom. Si l’objet NamedItem n’existe pas, la propriété isNull de l’objet renvoyé aura la valeur true.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getitemname-string"></a>getItem(name: string)
Obtient un objet NamedItem à l’aide de son nom.

#### <a name="syntax"></a>Syntaxe
```js
namedItemCollectionObject.getItem(name);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|name|string|nom de l’objet NamedItem.|

#### <a name="returns"></a>Retourne
[NamedItem](nameditem.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var sheetName = 'Sheet1';
    var nameditem = ctx.workbook.names.getItem(sheetName);
    nameditem.load('type');
    return ctx.sync().then(function() {
            console.log(nameditem.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### <a name="getitemornullname-string"></a>getItemOrNull(name: chaîne)
Obtient un objet NamedItem à l’aide de son nom. Si l’objet NamedItem n’existe pas, la propriété isNull de l’objet renvoyé aura la valeur true.

#### <a name="syntax"></a>Syntaxe
```js
namedItemCollectionObject.getItemOrNull(name);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|name|string|nom de l’objet NamedItem.|

#### <a name="returns"></a>Retourne
[NamedItem](nameditem.md)

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
    var nameditems = ctx.workbook.names;
    nameditems.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < nameditems.items.length; i++)
        {
            console.log(nameditems.items[i].name);
            console.log(nameditems.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


