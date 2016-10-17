# <a name="pagecontent-object-(javascript-api-for-onenote)"></a>Objet PageContent (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_  


Représente une zone sur une page qui contient des types de contenu de niveau supérieur tels que des plans ou des images. Un objet PageContent peut être affecté à une position XY.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|id|string|Obtient l’ID de l’objet PageContent. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-id)|
|left|double|Obtient ou définit la position à gauche (axe des abscisses) de l’objet PageContent.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-left)|
|top|double|Obtient ou définit la position supérieure (axe des ordonnées) de l’objet PageContent.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-top)|
|type|string|Obtient le type de l’objet PageContent. En lecture seule. Les valeurs possibles sont les suivantes : Outline, Image, Other.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-type)|

## <a name="relationships"></a>Relations
| Relation | Type   |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|image|[Image](image.md)|Obtient l’image dans l’objet PageContent. Renvoie une exception si PageContentType n’est pas défini sur Image. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-image)|
|ink|[FloatingInk](floatingink.md)|Obtient l’entrée manuscrite dans l’objet PageContent. Renvoie une exception si PageContentType n’est pas défini sur Ink. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-ink)|
|outline|[Outline](outline.md)|Obtient le plan de l’objet PageContent. Renvoie une exception si PageContentType n’est pas défini sur Outline. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-outline)|
|parentPage|[Page](page.md)|Obtient la page qui contient l’objet PageContent. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-parentPage)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[delete()](#delete)|void|Supprime l’objet PageContent.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-delete)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-load)|

## <a name="method-details"></a>Détails des méthodes


### <a name="delete()"></a>delete()
Supprime l’objet PageContent.

#### <a name="syntax"></a>Syntaxe
```js
pageContentObject.delete();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
OneNote.run(function (context) {

    var page = context.application.getActivePage();
    var pageContents = page.contents;

    var firstPageContent = pageContents.getItemAt(0);
    firstPageContent.load('type');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if(firstPageContent.isNull === false) {
                firstPageContent.delete();
                return context.sync();
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### <a name="load(param:-object)"></a>load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### <a name="syntax"></a>Syntaxe
```js
object.load(param);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### <a name="returns"></a>Retourne
void
