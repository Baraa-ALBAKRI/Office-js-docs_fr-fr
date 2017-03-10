# <a name="shape-object-javascript-api-for-visio"></a>Objet Shape (interface API JavaScript pour Visio)

S’applique à : _Visio Online_

Représente la classe Shape.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description|
|:---------------|:--------|:----------|
|id|int|Identificateur de l’objet Shape. En lecture seule.|
|name|chaîne|Nom de l’objet Shape. En lecture seule.|
|select|bool|Renvoie True si l’objet Shape est sélectionné. L’utilisateur peut le définir sur True pour sélectionner explicitement l’objet Shape.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-select)|
|text|chaîne|Texte de l’objet Shape. En lecture seule.|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type    |Description|
|:---------------|:--------|:----------|
|comments|[CommentCollection](commentcollection.md)|Renvoie la collection de commentaires. En lecture seule.|
|hyperlinks|[HyperlinkCollection](hyperlinkcollection.md)|Renvoie la collection Hyperlinks d’un objet Shape. En lecture seule.|
|shapeDataItems|[ShapeDataItemCollection](shapedataitemcollection.md)|Renvoie la section de données de l’objet Shape. En lecture seule.|
|subShapes|[ShapeCollection](shapecollection.md)|Obtient la collection SubShape. En lecture seule.|
|view|[ShapeView](shapeview.md)|Renvoie l’affichage de la forme. En lecture seule.|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[getBounds()](#getbounds)|[BoundingBox](boundingbox.md)|Renvoie l’objet BoundingBox qui spécifie le cadre englobant de la forme.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## <a name="method-details"></a>Détails des méthodes


### <a name="getbounds"></a>getBounds()
Renvoie l’objet BoundingBox qui spécifie le cadre englobant de la forme.

#### <a name="syntax"></a>Syntaxe
```js
shapeObject.getBounds();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Renvoie
[BoundingBox](boundingbox.md)

### <a name="loadparam-object"></a>load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### <a name="syntax"></a>Syntaxe
```js
object.load(param);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### <a name="returns"></a>Renvoie
void
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shapeName = "Sample Name";
    var shape = activePage.shapes.getItem(shapeName);
    shape.load();
    return ctx.sync().then(function () {
        console.log(shape.name );
        console.log(shape.id );
        console.log(shape.Text );
        console.log(shape.Select );
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
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
    shape.view.highlight = { color: "#E7E7E7", width: 100 };
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
