# <a name="highlight-object-javascript-api-for-visio"></a>Objet Highlight (API JavaScript pour Visio)

S’applique à : _Visio Online_

Représente les données mises en surbrillance ajoutées à la forme.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description|
|:---------------|:--------|:----------|
|color|string|Chaîne indiquant la couleur de la mise en surbrillance. Elle doit être au format « #RRGGBB », où chaque lettre représente un chiffre hexadécimal compris entre 0 et F, et où RR est la valeur rouge comprise entre 0 et 0xFF (255), GG la valeur verte comprise entre 0 et 0xFF (255), et BB la valeur bleue comprise entre 0 et 0xFF (255).|
|width|int|Nombre entier positif indiquant la largeur du trait de la mise en surbrillance en pixels.|

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
    shape.view.highlight = { color: "#E7E7E7", width: 100 };
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
