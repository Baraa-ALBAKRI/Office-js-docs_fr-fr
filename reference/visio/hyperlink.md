# <a name="hyperlink-object-javascript-api-for-visio"></a>Objet Hyperlink (interface API JavaScript pour Visio)

S’applique à : _Visio Online_

Représente l’objet Hyperlink.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description|
|:---------------|:--------|:----------|
|address|chaîne|Obtient l’adresse de l’objet Hyperlink. En lecture seule.|
|description|chaîne|Obtient la description d’un lien hypertexte. En lecture seule.|
|subAddress|chaîne|Obtient la sous-adresse de l’objet Hyperlink. En lecture seule.|

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
    var hyperlink = shape.hyperlinks.getItem(0);
    hyperlink.load();
    return ctx.sync().then(function() {
        console.log(hyperlink.description);
        console.log(hyperlink.address);
        console.log(hyperlink.subAddress);
     });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
