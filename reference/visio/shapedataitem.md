# <a name="shapedataitem-object-javascript-api-for-visio"></a>Objet ShapeDataItem (interface API JavaScript pour Visio)

S’applique à : _Visio Online_

Représente l’objet ShapeDataItem.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description|
|:---------------|:--------|:----------|
|format|string|Chaîne spécifiant le format de l’élément de données de forme. En lecture seule.|
|formattedValue|string|Chaîne spécifiant la valeur formatée de l’élément de données de forme. En lecture seule.|
|label|chaîne|Chaîne qui spécifie l’étiquette de l’élément de données de forme. En lecture seule.|
|value|chaîne|Chaîne qui spécifie la valeur de l’élément de données de forme. En lecture seule.|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## <a name="method-details"></a>Détails des méthodes


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
    var shape = activePage.shapes.getItem(0);
        var shapeDataItem = shape.shapeDataItems.getItem(0);
    shapeDataItem.load();
        return ctx.sync().then(function() {
                console.log(shapeDataItem.label);
                console.log(shapeDataItem.value);
        });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
