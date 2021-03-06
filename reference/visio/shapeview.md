# <a name="shapeview-object-javascript-api-for-visio"></a>Objet ShapeView (interface API JavaScript pour Visio)

S’applique à : _Visio Online_

Représente la classe ShapeView.

## <a name="properties"></a>Propriétés

Aucun

## <a name="relationships"></a>Relations
| Relation | Type    |Description|
|:---------------|:--------|:----------|
|highlight|[Highlight](highlight.md)|Représente la mise en surbrillance autour de la forme.|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[addOverlay(OverlayType: OverlayType, Content: chaîne, HorizontalAlignment: HorizontalAlignment, VerticalAlignment: VerticalAlignment, Width: valeur numérique, Height: valeur numérique)](#addoverlayoverlaytype-overlaytype-content-string-horizontalalignment-horizontalalignment-verticalalignment-verticalalignment-width-number-height-number)|int|Ajoute une superposition sur la forme.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|
|[removeOverlay(OverlayId: number)](#removeoverlayoverlayid-number)|void|Supprime une ou toutes les superpositions de la forme.|

## <a name="method-details"></a>Détails des méthodes


### <a name="addoverlayoverlaytype-overlaytype-content-string-horizontalalignment-horizontalalignment-verticalalignment-verticalalignment-width-number-height-number"></a>addOverlay(OverlayType: OverlayType, Content: chaîne, HorizontalAlignment: HorizontalAlignment, VerticalAlignment: VerticalAlignment, Width: valeur numérique, Height: valeur numérique)
Ajoute une superposition sur la forme.

#### <a name="syntax"></a>Syntaxe
```js
shapeViewObject.addOverlay(OverlayType, Content, HorizontalAlignment, VerticalAlignment, Width, Height);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|OverlayType|OverlayType|Type de superposition (texte, image).|
|Contenu|chaîne|Contenu de la superposition.|
|HorizontalAlignment|HorizontalAlignment|Alignement horizontal de la superposition (gauche, centre, droite)|
|VerticalAlignment|VerticalAlignment|Alignement vertical de la superposition (haut, milieu, bas)|
|Width|valeur numérique|Largeur de la superposition.|
|Height|valeur numérique|Hauteur de la superposition.|

#### <a name="returns"></a>Renvoie
int

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

### <a name="removeoverlayoverlayid-number"></a>removeOverlay(OverlayId: valeur numérique)
Supprime une ou toutes les superpositions de la forme.

#### <a name="syntax"></a>Syntaxe
```js
shapeViewObject.removeOverlay(OverlayId);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|OverlayId|number|ID de la superposition. Supprime l’ID de la superposition de la forme.|

#### <a name="returns"></a>Retourne
void
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

### <a name="property-access-examples"></a>Exemples d’accès aux propriétés
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
    var overlayId=shape.view.addOverlay(1, "Visio Online", 2, 2, 50, 50);
    return ctx.sync();
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
    shape.view.removeOverlay(1);
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
