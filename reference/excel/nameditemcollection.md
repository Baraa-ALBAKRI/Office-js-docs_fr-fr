# <a name="nameditemcollection-object-javascript-api-for-excel"></a>Objet NamedItemCollection (API JavaScript pour Excel)

Collection de tous les objets nameditem qui font partie du classeur ou de la feuille de calcul, en fonction de la méthode d’appel.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|éléments|[NamedItem[]](nameditem.md)|Collection d’objets NamedItem. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[add(name: string, reference: range ou string, comment: string)](#addname-string-reference-range-or-string-comment-string)|[NamedItem](nameditem.md)|Ajoute un nouveau nom à la collection de l’étendue donnée.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[addFormulaLocal(name: string, formula: string, comment: string)](#addformulalocalname-string-formula-string-comment-string)|[NamedItem](nameditem.md)|Ajoute un nouveau nom à la collection de l’étendue donnée à l’aide des paramètres régionaux de l’utilisateur pour la formule.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Obtient le nombre d’éléments nommés dans la collection.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|Obtient un objet NamedItem à l’aide de son nom.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[NamedItem](nameditem.md)|Obtient un objet nameditem à l’aide de son nom. Si l’objet nameditem n’existe pas, renvoie un objet null.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="addname-string-reference-range-or-string-comment-string"></a>add(name: string, reference: range ou string, comment: string)
Ajoute un nouveau nom à la collection de l’étendue donnée.

#### <a name="syntax"></a>Syntaxe
```js
namedItemCollectionObject.add(name, reference, comment);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|name|string|Nom de l’élément nommé.|
|reference|range ou string|Formule ou plage à laquelle le nom fait référence.|
|comment|string|Facultatif. Commentaire associé à l’élément nommé|

#### <a name="returns"></a>Renvoie
[NamedItem](nameditem.md)

### <a name="addformulalocalname-string-formula-string-comment-string"></a>addFormulaLocal(name: string, formula: string, comment: string)
Ajoute un nouveau nom à la collection de l’étendue donnée à l’aide des paramètres régionaux de l’utilisateur pour la formule.

#### <a name="syntax"></a>Syntaxe
```js
namedItemCollectionObject.addFormulaLocal(name, formula, comment);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|name|string|« Nom » de l’élément nommé.|
|formula|string|Formule dans les paramètres régionaux de l’utilisateur à laquelle le nom fait référence.|
|comment|string|Facultatif. Commentaire associé à l’élément nommé|

#### <a name="returns"></a>Renvoie
[NamedItem](nameditem.md)

### <a name="getcount"></a>getCount()
Obtient le nombre d’éléments nommés dans la collection.

#### <a name="syntax"></a>Syntaxe
```js
namedItemCollectionObject.getCount();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Renvoie
int

### <a name="getitemname-string"></a>getItem(name: string)
Obtient un objet NamedItem à l’aide de son nom.

#### <a name="syntax"></a>Syntaxe
```js
namedItemCollectionObject.getItem(name);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|name|string|nom de l’objet NamedItem.|

#### <a name="returns"></a>Retourne
[NamedItem](nameditem.md)

#### <a name="examples"></a>Exemples

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
### <a name="getitemornullobjectname-string"></a>getItemOrNullObject(name: string)
Obtient un objet nameditem à l’aide de son nom. Si l’objet nameditem n’existe pas, renvoie un objet null.

#### <a name="syntax"></a>Syntaxe
```js
namedItemCollectionObject.getItemOrNullObject(name);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|name|string|nom de l’objet NamedItem.|

#### <a name="returns"></a>Retourne
[NamedItem](nameditem.md)
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


