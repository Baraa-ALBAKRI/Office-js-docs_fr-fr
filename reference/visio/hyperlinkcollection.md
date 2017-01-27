# <a name="hyperlinkcollection-object-javascript-api-for-visio"></a>Objet HyperlinkCollection (interface API JavaScript pour Visio)

S’applique à : _Visio Online_
>**Remarque :** les API JavaScript Visio ne sont actuellement pas prises en charge dans les environnements d’évaluation ou de production.

Représente l’objet HyperlinkCollection.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description| Commentaires|
|:---------------|:--------|:----------|:---|
|items|[Hyperlink[]](hyperlink.md)|Collection d’objets de lien hypertexte. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-items)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:---|
|[getCount()](#getcount)|int|Obtient le nombre de liens hypertexte.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-getCount)|
|[getItem(Key: valeur numérique ou chaîne)](#getitemkey-number-or-string)|[Hyperlink](hyperlink.md)|Obtient un objet Hyperlink à l’aide de sa clé (nom ou ID).|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-getItem)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-load)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getcount"></a>getCount()
Obtient le nombre de liens hypertexte.

#### <a name="syntax"></a>Syntaxe
```js
hyperlinkCollectionObject.getCount();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Renvoie
int

### <a name="getitemkey-number-or-string"></a>getItem(Key: valeur numérique ou chaîne)
Obtient un objet Hyperlink à l’aide de sa clé (nom ou ID).

#### <a name="syntax"></a>Syntaxe
```js
hyperlinkCollectionObject.getItem(Key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|Clé|valeur numérique ou chaîne|La clé est le nom ou l’index de l’objet Hyperlink à récupérer.|

#### <a name="returns"></a>Renvoie
[Hyperlink](hyperlink.md)

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
    var shapeName = "Manager Belt";
    var shape = activePage.shapes.getItem(shapeName);
    var hyperlinks = shape.hyperlinks;
    shapeHyperlinks.load();
        ctx.sync().then(function () {
            for(var i=0; i<shapeHyperlinks.items.length;i++)
                {
                  var hyperlink = shapeHyperlinks.items[i];
                  console.log("Description:"+hyperlink.description +"Address:"+hyperlink.address +"SubAddress:  "+ hyperlink.subAddress);
                }

            });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
