# <a name="shapedataitemcollection-object-javascript-api-for-visio"></a>Objet ShapeDataItemCollection (interface API JavaScript pour Visio)

S’applique à : _Visio Online_
>**Remarque :** les API JavaScript Visio ne sont actuellement pas prises en charge dans les environnements d’évaluation ou de production.

Représente la collection d’éléments de données d’une forme donnée.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description| Commentaires|
|:---------------|:--------|:----------|:---|
|items|[ShapeDataItem[]](shapedataitem.md)|Collection d’objets ShapeDataItem. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItemCollection-items)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:---|
|[getCount()](#getcount)|int|Obtient le nombre d’éléments de données de forme.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItemCollection-getCount)|
|[getItem(key: chaîne)](#getitemkey-string)|[ShapeDataItem](shapedataitem.md)|Obtient l’élément de données de forme à l’aide de son nom.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItemCollection-getItem)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItemCollection-load)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getcount"></a>getCount()
Obtient le nombre d’éléments de données de forme.

#### <a name="syntax"></a>Syntaxe
```js
shapeDataItemCollectionObject.getCount();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Renvoie
int

### <a name="getitemkey-string"></a>getItem(key: chaîne)
Obtient l’élément de données de forme à l’aide de son nom.

#### <a name="syntax"></a>Syntaxe
```js
shapeDataItemCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|Key|chaîne|La clé est le nom de l’élément de données de forme à récupérer.|

#### <a name="returns"></a>Renvoie
[ShapeDataItem](shapedataitem.md)

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

#### <a name="returns"></a>Renvoie
void
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
        var shapeDataItems = shape.shapeDataItems;
        shapeDataItems.load();
        return ctx.sync().then(function() {
            for (var i = 0; i < shapeDataItems.items.length; i++)
            {
                console.log(shapeDataItems.items[i].label);
                console.log(shapeDataItems.items[i].value);
            }
        });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
