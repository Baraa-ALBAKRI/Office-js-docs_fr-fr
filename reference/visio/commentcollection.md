# <a name="commentcollection-object-javascript-api-for-visio"></a>Objet CommentCollection (API JavaScript pour Visio)

S’applique à : _Visio Online_

Représente la collection de commentaires d’une forme donnée.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description
|:---------------|:--------|:----------|
|items|[Comment[]](comment.md)|Collection des objets de commentaire. En lecture seule.|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[getCount()](#getcount)|int|Obtient le nombre de commentaires.|
|[getItem(key: string)](#getitemkey-string)|[Comment](comment.md)|Obtient le commentaire à l’aide de son nom.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## <a name="method-details"></a>Détails des méthodes


### <a name="getcount"></a>getCount()
Obtient le nombre de commentaires.

#### <a name="syntax"></a>Syntaxe
```js
CommentCollectionObject.getCount();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Renvoie
int

### <a name="getitemkey-string"></a>getItem(key: string)
Obtient le commentaire à l’aide de son nom.

#### <a name="syntax"></a>Syntaxe
```js
CommentCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|Key|string|« Key » (clé) est le nom du commentaire à récupérer.|

#### <a name="returns"></a>Renvoie
[Comment](comment.md)

### <a name="loadparam-object"></a>load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### <a name="syntax"></a>Syntaxe
```js
object.load(param);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### <a name="returns"></a>Renvoie
void
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés
```js
 Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shapeName = "Position Belt.41";
    var shape = activePage.shapes.getItem(shapeName);
    var shapecomments= shape.comments;
        shapecomments.load();
        return ctx.sync().then(function () {
             for(var i=0; i<shapecomments.items.length;i++)
        {
                    var comment= shapecomments.items[i];
            console.log("comment Author: " + comment.author);
            console.log("Comment Text: " + comment.text);
            console.log("Date " + comment.date);
        }
     });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
