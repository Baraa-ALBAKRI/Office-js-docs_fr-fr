# <a name="pageview-object-javascript-api-for-visio"></a>Objet PageView (interface API JavaScript pour Visio)

S’applique à : _Visio Online_

Représente la classe PageView.

## <a name="properties"></a>Propriétés

| Propriété | Type |Description|
|:---------------|:--------|:----------|
|zoom|int|Obtient/Définit le niveau de zoom de l’objet Page.|

## <a name="relationships"></a>Relations
Aucun

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[centerViewportOnShape(ShapeId: number)](#centerviewportonshapeshapeid-number)|void|Effectue un panoramique du dessin Visio pour placer la forme spécifiée au centre de l’affichage.|
|[fitToWindow()](#fittowindow)|void|Ajuste l’objet Page à la fenêtre active.|
|[getPosition()](#getposition)|[Position](position.md)|Spécifie la position de la page affichée.|
|[getSelection()](#getselection)|[Selection](selection.md)|Représente la sélection dans la page.|
|[isShapeInViewport(Shape: Shape)](#isshapeinviewportshape-shape)|bool|Vérifie si la forme se situe devant l’objet Page ou non.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|
|[setPosition(Position: Position)](#setpositionposition-position)|void|Définit la position de la page dans l’affichage.|

## <a name="method-details"></a>Détails des méthodes


### <a name="centerviewportonshapeshapeid-number"></a>centerViewportOnShape(ShapeId: valeur numérique)
Effectue un panoramique du dessin Visio pour placer la forme spécifiée au centre de l’affichage.

#### <a name="syntax"></a>Syntaxe
```js
pageViewObject.centerViewportOnShape(ShapeId);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|ShapeId|valeur numérique|Affiche ShapeId au centre.|

#### <a name="returns"></a>Renvoie
void

#### <a name="examples"></a>Exemples
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
    activePage.view.centerViewportOnShape(shape.Id);
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="fittowindow"></a>fitToWindow()
Ajuste l’objet Page à la fenêtre active.

#### <a name="syntax"></a>Syntaxe
```js
pageViewObject.fitToWindow();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

### <a name="getposition"></a>getPosition()
Spécifie la position de la page affichée.

#### <a name="syntax"></a>Syntaxe
```js
pageViewObject.getPosition();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Renvoie
[Position](position.md)

### <a name="getselection"></a>getSelection()
Représente la sélection dans la page.

#### <a name="syntax"></a>Syntaxe
```js
pageViewObject.getSelection();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Renvoie
[Selection](selection.md)

### <a name="isshapeinviewportshape-shape"></a>isShapeInViewport(Shape: Shape)
Vérifie si la forme se situe devant l’objet Page ou non.

#### <a name="syntax"></a>Syntaxe
```js
pageViewObject.isShapeInViewport(Shape);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|Shape|Shape|Forme à vérifier.|

#### <a name="returns"></a>Renvoie
bool

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

### <a name="setpositionposition-position"></a>setPosition(Position: Position)
Définit la position de la page dans l’affichage.

#### <a name="syntax"></a>Syntaxe
```js
pageViewObject.setPosition(Position);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|Position|Position|Spécifie la nouvelle position de la page affichée.|

#### <a name="returns"></a>Renvoie
void
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    activePage.view.zoom = 300;
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

