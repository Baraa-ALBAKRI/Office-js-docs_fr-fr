# <a name="outline-object-(javascript-api-for-onenote)"></a>Objet Outline (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_  


Représente un conteneur pour les objets Paragraph.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|id|chaîne|Obtient l’ID de l’objet Outline. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-id)|

## <a name="relationships"></a>Relations
| Relation | Type   |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|pageContent|[PageContent](pagecontent.md)|Obtient l’objet PageContent qui contient le plan. Cet objet définit la position du plan sur la page. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-pageContent)|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Obtient la collection d’objets Paragraph dans le plan. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-paragraphs)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[appendHtml(html: string)](#appendhtmlhtml-string)|void|Ajoute le code HTML spécifié dans la partie inférieure du plan.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendHtml)|
|[appendImage(base64EncodedImage: string, width: double, height: double)](#appendimagebase64encodedimage-string-width-double-height-double)|[Image](image.md)|Ajoute l’image spécifiée dans la partie inférieure du plan.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendImage)|
|[appendRichText(paragraphText: string)](#appendrichtextparagraphtext-string)|[RichText](richtext.md)|Ajoute le texte spécifié dans la partie inférieure du plan.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendRichText)|
|[appendTable(rowCount: number, columnCount: number, values: string[][])](#appendtablerowcount-number-columncount-number-values-string)|[Table](table.md)|Ajoute un tableau avec le nombre spécifié de lignes et de colonnes dans la partie inférieure du plan.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendTable)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-load)|

## <a name="method-details"></a>Détails des méthodes


### <a name="appendhtml(html:-string)"></a>appendHtml(html: string)
Ajoute le code HTML spécifié dans la partie inférieure du plan.

#### <a name="syntax"></a>Syntaxe
```js
outlineObject.appendHtml(html);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|Html|chaîne|Chaîne HTML à ajouter. Voir [HTML pris en charge](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) pour l’API JavaScript des compléments OneNote.|

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Get pageContents of the activePage. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            if (pageContents.items.length != 0 && pageContents.items[0].type == "Outline")
            {
                // First item is an outline.
                outline = pageContents.items[0].outline;

                // Queue a command to append a paragraph to the outline.
                outline.appendHtml("<p>new paragraph</p>");

                // Run the queued commands.
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


### <a name="appendimage(base64encodedimage:-string,-width:-double,-height:-double)"></a>appendImage(base64EncodedImage: string, width: double, height: double)
Ajoute l’image spécifiée dans la partie inférieure du plan.

#### <a name="syntax"></a>Syntaxe
```js
outlineObject.appendImage(base64EncodedImage, width, height);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Chaîne HTML à ajouter.|
|width|double|Facultatif. Largeur de l’unité des points. La valeur par défaut est Null et la largeur d’image est respectée.|
|height|double|Facultatif. Hauteur de l’unité des points. La valeur par défaut est Null et la hauteur d’image est respectée.|

#### <a name="returns"></a>Retourne
[Image](image.md)

### <a name="appendrichtext(paragraphtext:-string)"></a>appendRichText(paragraphText: string)
Ajoute le texte spécifié dans la partie inférieure du plan.

#### <a name="syntax"></a>Syntaxe
```js
outlineObject.appendRichText(paragraphText);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|paragraphText|string|Chaîne HTML à ajouter.|

#### <a name="returns"></a>Retourne
[RichText](richtext.md)

### <a name="appendtable(rowcount:-number,-columncount:-number,-values:-string[][])"></a>appendTable(rowCount: number, columnCount: number, values: string[][])
Ajoute un tableau avec le nombre spécifié de lignes et de colonnes dans la partie inférieure du plan.

#### <a name="syntax"></a>Syntaxe
```js
outlineObject.appendTable(rowCount, columnCount, values);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|rowCount|number|Obligatoire. Nombre de lignes dans le tableau.|
|columnCount|number|Obligatoire. Nombre de colonnes dans le tableau.|
|values|string[][]|Facultatif. Tableau 2D facultatif. Les cellules sont remplies si les chaînes correspondantes sont spécifiées dans le tableau.|

#### <a name="returns"></a>Retourne
[Table](table.md)

#### <a name="examples"></a>Exemples
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Get pageContents of the activePage. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            if (pageContents.items.length != 0 && pageContents.items[0].type == "Outline") {
                // First item is an outline.
                var outline = pageContents.items[0].outline;

                // Queue a command to append a paragraph to the outline.
                outline.appendTable(2, 2, [[1, 2],[3, 4]]);

                // Run the queued commands.
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
