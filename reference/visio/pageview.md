# <a name="pageview-object-javascript-api-for-visio"></a>Objet PageView (interface API JavaScript pour Visio)

S’applique à : _Visio Online_
>**Remarque :** Les interfaces API JavaScript pour Visio sont actuellement affichées dans l’aperçu et peuvent être modifiées. Elles ne sont actuellement pas prises en charge dans les environnements de production.

Représente la classe PageView.

## <a name="properties"></a>Propriétés

| Propriété | Type |Description| Commentaires|
|:---------------|:--------|:----------|:---|
|zoom|int|Obtient/Définit le niveau de zoom de l’objet Page.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-zoom)|

## <a name="relationships"></a>Relations

Aucun

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:---|
|[centerViewportOnShape(ShapeId: valeur numérique)](#centerviewportonshapeshapeid-number)|void|Effectue un panoramique du dessin Visio pour placer la forme spécifiée au centre de l’affichage.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-centerViewportOnShape)|
|[fitToWindow()](#fittowindow)|void|Ajuste l’objet Page à la fenêtre active.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-fitToWindow)|
|[isShapeInViewport(Shape: Shape)](#isshapeinviewportshape-shape)|bool|Vérifie si la forme se situe devant l’objet Page ou non.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-isShapeInViewport)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-load)|

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

