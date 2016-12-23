# <a name="shapecollection-object-javascript-api-for-visio"></a>Objet ShapeCollection (interface API JavaScript pour Visio)

S’applique à : _Visio Online_
>**Remarque :** Les interfaces API JavaScript pour Visio sont actuellement affichées dans l’aperçu et peuvent être modifiées. Elles ne sont actuellement pas prises en charge dans les environnements de production.

Représente la collection Shape.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description| Commentaires|
|:---------------|:--------|:----------|:---|
|items|[Shape[]](shape.md)|Collection d’objets de forme. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-items)|

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:---|
|[getCount()](#getcount)|int|Obtient le nombre de formes de la collection.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-getCount)|
|[getItem(key: valeur numérique ou chaîne)](#getitemkey-number-or-string)|[Shape](shape.md)|Obtient une forme à l’aide de sa clé (nom ou index).|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-getItem)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-load)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getcount"></a>getCount()
Obtient le nombre de formes de la collection.

#### <a name="syntax"></a>Syntaxe
```js
shapeCollectionObject.getCount();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Renvoie
int

#### <a name="examples"></a>Exemples
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var numShapesActivePage = activePage.shapes.getCount();
    return ctx.sync().then(function () {
        console.log("Shapes Count: " + numShapesActivePage.value);
    });

}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getitemkey-number-or-string"></a>getItem(key: valeur numérique ou chaîne)
Obtient une forme à l’aide de sa clé (nom ou index).

#### <a name="syntax"></a>Syntaxe
```js
shapeCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|Key|valeur numérique ou chaîne|La clé est le nom ou l’index de la forme à récupérer.|

#### <a name="returns"></a>Renvoie
[Shape](shape.md)

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
