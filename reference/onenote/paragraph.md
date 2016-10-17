# <a name="paragraph-object-(javascript-api-for-onenote)"></a>Objet Paragraph (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_  


Conteneur pour le contenu visible d’une page. Un paragraphe peut contenir n’importe quel type de contenu ParagraphType.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|id|chaîne|Obtient l’ID de l’objet Paragraph. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-id)|
|type|string|Obtient le type de l’objet Paragraph. En lecture seule. Les valeurs possibles sont les suivantes : RichText, Image, Table, Other.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-type)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type   |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|image|[Image](image.md)|Renvoie l’objet Image dans le paragraphe. Renvoie une exception si ParagraphType n’est pas défini sur Image. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-image)|
|inkWords|[InkWordCollection](inkwordcollection.md)|Obtient la collection Ink dans le paragraphe. Renvoie une exception si ParagraphType n’est pas défini sur Ink. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-inkWords)|
|outline|[Outline](outline.md)|Renvoie l’objet Outline qui contient le paragraphe. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-outline)|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Collection de paragraphes sous ce paragraphe. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-paragraphs)|
|parentParagraph|[Paragraph](paragraph.md)|Obtient l’objet de paragraphe parent. Indique si un paragraphe parent n’existe pas. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentParagraph)|
|parentParagraphOrNull|[Paragraph](paragraph.md)|Obtient l’objet de paragraphe parent. Renvoie la valeur null si un paragraphe parent n’existe pas. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentParagraphOrNull)|
|parentTableCell|[TableCell](tablecell.md)|Obtient l’objet TableCell qui contient le paragraphe s’il en existe un. Si le parent n’est pas un objet TableCell, renvoie ItemNotFound. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentTableCell)|
|parentTableCellOrNull|[TableCell](tablecell.md)|Obtient l’objet TableCell qui contient le paragraphe s’il en existe un. Si le parent n’est pas un objet TableCell, renvoie la valeur null. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentTableCellOrNull)|
|richText|[RichText](richtext.md)|Renvoie l’objet RichText du paragraphe. Renvoie une exception si ParagraphType n’est pas défini sur RichText. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-richText)|
|table|[Table](table.md)|Obtient l’objet Table dans le paragraphe. Renvoie une exception si ParagraphType n’est pas défini sur Table. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-table)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[delete()](#delete)|void|Supprime le paragraphe.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-delete)|
|[insertHtmlAsSibling(insertLocation: string, html: string)](#inserthtmlassiblinginsertlocation-string-html-string)|void|Insère le contenu HTML spécifié.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertHtmlAsSibling)|
|[insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: double, height: double)](#insertimageassiblinginsertlocation-string-base64encodedimage-string-width-double-height-double)|[Image](image.md)|Insère l’image à l’emplacement d’insertion spécifié.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertImageAsSibling)|
|[insertRichTextAsSibling(insertLocation: string, paragraphText: string)](#insertrichtextassiblinginsertlocation-string-paragraphtext-string)|[RichText](richtext.md)|Insère le texte du paragraphe à l’emplacement d’insertion spécifié.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertRichTextAsSibling)|
|[insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values: string[][])](#inserttableassiblinginsertlocation-string-rowcount-number-columncount-number-values-string)|[Table](table.md)|Ajoute un tableau avec le nombre spécifié de lignes et de colonnes avant ou après le paragraphe en cours.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertTableAsSibling)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-load)|

## <a name="method-details"></a>Détails des méthodes


### <a name="delete()"></a>delete()
Supprime le paragraphe.

#### <a name="syntax"></a>Syntaxe
```js
paragraphObject.delete();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    
    var paragraphs = pageContent.outline.paragraphs;
    
    var firstParagraph = paragraphs.getItemAt(0);
    
    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Queue a command to delete the first paragraph                 
            firstParagraph.delete();
            
            // Run the command to delete it
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="inserthtmlassibling(insertlocation:-string,-html:-string)"></a>insertHtmlAsSibling(insertLocation: string, html: string)
Insère le contenu HTML spécifié.

#### <a name="syntax"></a>Syntaxe
```js
paragraphObject.insertHtmlAsSibling(insertLocation, html);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|insertLocation|string|Emplacement du nouveau contenu relatif au paragraphe actif.  Les valeurs possibles sont les suivantes : Before, After|
|Html|string|Chaîne HTML qui décrit la présentation visuelle du contenu. Voir [HTML pris en charge](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) pour l’API JavaScript des compléments OneNote.|

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    var paragraphs = pageContent.outline.paragraphs;
    var firstParagraph = paragraphs.getItemAt(0);

    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Queue commands to insert before and after the first paragraph
            firstParagraph.insertHtmlAsSibling("Before", "<p>ContentBeforeFirstParagraph</p>");
            firstParagraph.insertHtmlAsSibling("After", "<p>ContentAfterFirstParagraph</p>");
            
            // Run the command to run inserts
            return context.sync();
        });
))
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="insertimageassibling(insertlocation:-string,-base64encodedimage:-string,-width:-double,-height:-double)"></a>insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: double, height: double)
Insère l’image à l’emplacement d’insertion spécifié.

#### <a name="syntax"></a>Syntaxe
```js
paragraphObject.insertImageAsSibling(insertLocation, base64EncodedImage, width, height);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|insertLocation|chaîne|Emplacement du tableau relatif au paragraphe actif.  Les valeurs possibles sont les suivantes : Before, After|
|base64EncodedImage|string|Chaîne HTML à ajouter.|
|width|double|Facultatif. Largeur de l’unité des points. La valeur par défaut est Null et la largeur d’image est respectée.|
|height|double|Facultatif. Hauteur de l’unité des points. La valeur par défaut est Null et la hauteur d’image est respectée.|

#### <a name="returns"></a>Retourne
[Image](image.md)

#### <a name="examples"></a>Exemples
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    var paragraphs = pageContent.outline.paragraphs;
    var firstParagraph = paragraphs.getItemAt(0);

    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Queue commands to insert before and after the first paragraph
            firstParagraph.insertImageAsSibling("Before", "R0lGODlhDwAPAKECAAAAzMzM/////wAAACwAAAAADwAPAAACIISPeQHsrZ5ModrLlN48CXF8m2iQ3YmmKqVlRtW4MLwWACH+H09wdGltaXplZCBieSBVbGVhZCBTbWFydFNhdmVyIQAAOw==");
            firstParagraph.insertImageAsSibling("After", "R0lGODlhDwAPAKECAAAAzMzM/////wAAACwAAAAADwAPAAACIISPeQHsrZ5ModrLlN48CXF8m2iQ3YmmKqVlRtW4MLwWACH+H09wdGltaXplZCBieSBVbGVhZCBTbWFydFNhdmVyIQAAOw==");
            
            // Run the command to insert images
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="insertrichtextassibling(insertlocation:-string,-paragraphtext:-string)"></a>insertRichTextAsSibling(insertLocation: string, paragraphText: string)
Insère le texte du paragraphe à l’emplacement d’insertion spécifié.

#### <a name="syntax"></a>Syntaxe
```js
paragraphObject.insertRichTextAsSibling(insertLocation, paragraphText);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|insertLocation|chaîne|Emplacement du tableau relatif au paragraphe actif.  Les valeurs possibles sont les suivantes : Before, After|
|paragraphText|string|Chaîne HTML à ajouter.|

#### <a name="returns"></a>Retourne
[RichText](richtext.md)

#### <a name="examples"></a>Exemples
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    var paragraphs = pageContent.outline.paragraphs;
    var firstParagraph = paragraphs.getItemAt(0);

    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Queue commands to insert before and after the first paragraph
            firstParagraph.insertRichTextAsSibling("Before", "Text Appears Before Paragraph");
            firstParagraph.insertRichTextAsSibling("After", "Text Appears After Paragraph");
            
            // Run the command to insert text contents
            return context.sync();
        });
})  
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```


### <a name="inserttableassibling(insertlocation:-string,-rowcount:-number,-columncount:-number,-values:-string[][])"></a>insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values: string[][])
Ajoute un tableau avec le nombre spécifié de lignes et de colonnes avant ou après le paragraphe en cours.

#### <a name="syntax"></a>Syntaxe
```js
paragraphObject.insertTableAsSibling(insertLocation, rowCount, columnCount, values);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|insertLocation|chaîne|Emplacement du tableau relatif au paragraphe actif.  Les valeurs possibles sont les suivantes : Before, After|
|rowCount|number|Nombre de lignes dans le tableau.|
|columnCount|number|Nombre de colonnes dans le tableau.|
|values|string[][]|Facultatif. Tableau 2D facultatif. Les cellules sont remplies si les chaînes correspondantes sont spécifiées dans le tableau.|

#### <a name="returns"></a>Retourne
[Table](table.md)

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

**id et type**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;
    
    // Queue a command to load the outline property of each pageContent.
    pageContents.load("outline");
        
    // Get the first PageContent on the page, and then get its Outline.
    var pageContent = pageContents._GetItem(0);
    var paragraphs = pageContent.outline.paragraphs;
            
    // Queue a command to load the id and type of each paragraph.
    paragraphs.load("id,type");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Write the text.                  
            $.each(paragraphs.items, function(index, paragraph) {
                console.log("Paragraph type: " + paragraph.type);
                console.log("Paragraph ID: " + paragraph.id);
            });
        });
})      
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```

**paragraphs**
```js
OneNote.run(function(context) {
    var app = context.application;
    
    // Gets the active outline
    var outline = app.getActiveOutline();
    
    // load nested paragraphs and their types.
    outline.load("paragraphs/type");
    
    return context.sync().then(function () {
        var paragraphs = outline.paragraphs.items;
        
        var promise;
        // for each nested paragraphs, load tables only
        for (var i = 0; i < paragraphs.length; i++) {
            var paragraph = paragraphs[i];
            if (paragraph.type == "Table") {
                paragraph.load("table/id");
                promise =  context.sync().then(function() {
                    console.log(paragraph.table.id);
                });
            }
        }
        return promise;
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

