# <a name="floatingink-object-(javascript-api-for-onenote)"></a>Objet FloatingInk (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_  


Représente un groupe de traits d’encre.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|id|chaîne|Obtient l’ID de l’objet FloatingInk. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-id)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type   |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|inkStrokes|[InkStrokeCollection](inkstrokecollection.md)|Obtient les traits de l’objet FloatingInk. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-inkStrokes)|
|pageContent|[PageContent](pagecontent.md)|Obtient le parent PageContent de l’objet FloatingInk. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-pageContent)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-load)|

## <a name="method-details"></a>Détails des méthodes


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
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

**id**
```js
OneNote.run(function(context) {

    // Gets the active page.
    var page = context.application.getActivePage();
    var contents = page.contents;
    
    // Load page contents and their types.
    page.load('contents/type');
    return context.sync()
        .then(function(){
        
            // Load every ink content.
            $.each(contents.items, function(i, content) {
                if (content.type == "Ink")
                {
                    content.load('ink/id');
                }                           
            })
            return context.sync();
        })
        .then(function(){
        
            // Log ID of every ink content.
            $.each(contents.items, function(i, content) {
                if (content.type == "Ink")
                {
                    console.log(content.ink.id);
                }                           
            })              
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```
