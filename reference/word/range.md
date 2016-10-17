# <a name="range-object-(javascript-api-for-word)"></a>Objet Range (interface API JavaScript pour Word)

Représente une zone contiguë dans un document.

_S’applique à : Word 2016, Word pour iPad, Word pour Mac, Word Online_

## <a name="properties"></a>Propriétés
| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|style|string|Obtient ou définit le style utilisé pour la plage. Il s’agit du nom du style pré-installé ou personnalisé.|
|text|string|Obtient le texte de la plage. En lecture seule.|

## <a name="relationships"></a>Relations
| Relation | Type   |Description|
|:---------------|:--------|:----------|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Obtient la collection d’objets de contrôle de contenu qui se trouvent dans la plage. En lecture seule.|
|police|[Font](font.md)|Obtient le format de texte de la plage. Utilisez cette propriété pour obtenir et définir le nom de la police, la taille, la couleur et d’autres propriétés. En lecture seule.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Obtient la collection d’objets inlinePicture qui se trouvent dans la plage. En lecture seule.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Obtient la collection d’objets de paragraphe qui se trouvent dans la plage. En lecture seule.|
|parentContentControl|[ContentControl](contentcontrol.md)|Obtient le contrôle de contenu qui contient la plage. Renvoie null s’il n’existe pas de contrôle de contenu parent. En lecture seule.|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Efface le contenu de l’objet de plage. L’utilisateur peut effectuer l’opération d’annulation sur le contenu effacé.|
|[delete()](#delete)|void|Supprime la plage et son contenu du document.|
|[getHtml()](#gethtml)|chaîne|Obtient la représentation HTML de l’objet de plage.|
|[getOoxml()](#getooxml)|chaîne|Obtient la représentation OOXML de l’objet de plage.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Insère un saut à l’emplacement spécifié. Un saut peut uniquement être inséré dans des objets de plage qui sont contenus dans le corps de document principal, sauf s’il s’agit d’un saut de ligne, auquel cas il peut être inséré dans n’importe quel objet de corps. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Encadre l’objet de plage avec un contrôle de contenu de texte enrichi.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Insère un document dans la plage à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Insère du code HTML dans la plage à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Insère une image dans la plage à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Replace » (remplacer), « Start » (début), « End » (fin), « Before » (avant) ou « After » (après).
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Insère du code OOXML ou un élément wordProcessingML dans la plage, à l’emplacement spécifié.  La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Insère un paragraphe dans la plage à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Insère du texte dans la plage à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestringssearchoptions)|[SearchResultCollection](searchresultcollection.md)|Effectue une recherche avec les options de recherche spécifiées dans l’étendue de l’objet de plage. Les résultats de la recherche sont un ensemble d’objets de plage.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Sélectionne la plage et y accède via l’interface utilisateur de Word. Les valeurs selectionMode peuvent être « Select » (sélectionner), « Start » (début) ou « End » (fin).|

## <a name="method-details"></a>Détails de méthodes

### <a name="clear()"></a>clear()
Efface le contenu de l’objet de plage. L’utilisateur peut effectuer l’opération d’annulation sur le contenu effacé.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.clear();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to clear the contents of the proxy range object.
    range.clear();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Cleared the selection (range object)');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
### <a name="delete()"></a>delete()
Supprime la plage et son contenu du document.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.delete();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to delete the range object.
    range.delete();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Deleted the selection (range object)');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="gethtml()"></a>getHtml()
Obtient la représentation HTML de l’objet de plage.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getHtml();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
string

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to get the HTML of the current selection.
    var html = range.getHtml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The HTML read from the document was: ' + html.value);
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="getooxml()"></a>getOoxml()
Obtient la représentation OOXML de l’objet de plage.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.getOoxml();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
string

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to get the OOXML of the current selection.
    var ooxml = range.getOoxml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The OOXML read from the document was:  ' + ooxml.value);
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="insertbreak(breaktype:-breaktype,-insertlocation:-insertlocation)"></a>insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Insère un saut à l’emplacement spécifié. Un saut peut uniquement être inséré dans des objets de plage qui sont contenus dans le corps de document principal, sauf s’il s’agit d’un saut de ligne, auquel cas il peut être inséré dans n’importe quel objet de corps. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.insertBreak(breakType, insertLocation);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|breakType|BreakType|Obligatoire. Type de saut à ajouter à la plage.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Before » (avant) » ou « After » (après).|

#### <a name="returns"></a>Retourne
void

#### <a name="additional-details"></a>Détails supplémentaires
À l’exception des sauts de ligne, vous ne pouvez pas insérer de saut dans les objets d’en-tête, de pied de page, de note de bas de page, de note de fin, de commentaire et de zone de texte.

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert a page break after the selected text.
    range.insertBreak('page', 'After');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Inserted a page break after the selected text.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="insertcontentcontrol()"></a>insertContentControl()
Encadre l’objet de plage avec un contrôle de contenu de texte enrichi.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.insertContentControl();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[ContentControl](contentcontrol.md)

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert a content control around the selected text,
    // and create a proxy content control object. We'll update the properties
    // on the content control.
    var myContentControl = range.insertContentControl();
    myContentControl.tag = "Customer-Address";
    myContentControl.title = "Enter Customer Address Here:";
    myContentControl.style = "Normal";
    myContentControl.insertText("One Microsoft Way, Redmond, WA 98052", 'replace');
    myContentControl.cannotEdit = true;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Wrapped a content control around the selected text.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="insertfilefrombase64(base64file:-string,-insertlocation:-insertlocation)"></a>insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Insère un document dans la plage à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.insertFileFromBase64(base64File, insertLocation);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|base64File|string|Obligatoire. Contenu du fichier encodé au format Base64 à insérer.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert base64 encoded .docx at the beginning of the range.
    // You'll need to implement getBase64() to make this work.
    range.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Added base64 encoded text to the beginning of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="inserthtml(html:-string,-insertlocation:-insertlocation)"></a>insertHtml(html: string, insertLocation: InsertLocation)
Insère du code HTML dans la plage à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.insertHtml(html, insertLocation);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|Html|string|Obligatoire. Code HTML à insérer dans la plage.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert HTML in to the beginning of the range.
    range.insertHtml('<strong>This is text inserted with range.insertHtml()</strong>', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('HTML added to the beginning of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="insertinlinepicturefrombase64(base64encodedimage:-string,-insertlocation:-insertlocation)"></a>insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Insère une image dans la plage à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début), « End » (fin), « Before » (avant) ou « After » (après).

#### <a name="syntax"></a>Syntaxe
rangeObject.insertInlinePictureFromBase64(image, insertLocation);

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Obligatoire. Image encodée au format Base64 à insérer dans la plage.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Replace » (remplacer), « Start » (début), « End » (fin), « Before » (avant) ou « After » (après).|

#### <a name="returns"></a>Retourne
[InlinePicture](inlinepicture.md)

### <a name="insertooxml(ooxml:-string,-insertlocation:-insertlocation)"></a>insertOoxml(ooxml: string, insertLocation: InsertLocation)
Insère du contenu OOXML ou wordProcessingML dans la plage, à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.insertOoxml(ooxml, insertLocation);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|ooxml|string|Obligatoire. Contenu OOXML ou wordProcessingML à insérer dans la plage.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert OOXML in to the beginning of the range.
    range.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('OOXML added to the beginning of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### <a name="additional-information"></a>Informations supplémentaires
Pour obtenir des instructions sur l'utilisation d’OOXML, voir [Création de compléments plus performants pour Word avec Office Open XML](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx).

### <a name="insertparagraph(paragraphtext:-string,-insertlocation:-insertlocation)"></a>insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Insère un paragraphe dans la plage à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.insertParagraph(paragraphText, insertLocation);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|paragraphText|string|Obligatoire. Texte de paragraphe à insérer.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Before » (avant) » ou « After » (après).|

#### <a name="returns"></a>Retourne
[Paragraph](paragraph.md)

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert the paragraph after the range.
    range.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Paragraph added to the end of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="inserttext(text:-string,-insertlocation:-insertlocation)"></a>insertText(text: string, insertLocation: InsertLocation)
Insère du texte dans la plage à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.insertText(text, insertLocation);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|text|string|Obligatoire. Texte à insérer.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert the paragraph at the end of the range.
    range.insertText('New text inserted into the range.', Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Text added to the end of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
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

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to load font and style information for the range.
    context.load(range, 'font/size, font/name, font/color, style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Show the results of the load method. Here we show the
        // property values on the range object.
        var results = "  ---Font size: " + range.font.size +
                      "  ---Font name: " + range.font.name +
                      "  ---Font color: " + range.font.color +
                      "  ---Style: " + range.style;
        console.log(results);
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search(searchtext:-string,-searchoptions:-paramtypestrings.searchoptions)"></a>search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)
Effectue une recherche avec les options de recherche spécifiées dans l’étendue de l’objet de plage. Les résultats de la recherche sont un ensemble d’objets de plage.

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.search(searchText, searchOptions);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|searchText|string|Obligatoire. Texte de recherche.|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|Facultatif. Options de la recherche.|

#### <a name="returns"></a>Retourne
[SearchResultCollection](searchresultcollection.md)


### <a name="select(selectionmode:-selectionmode)"></a>select(selectionMode: SelectionMode)
Sélectionne la plage et y accède via l’interface utilisateur de Word. Les valeurs selectionMode peuvent être « Select » (sélectionner), « Start » (début) ou « End » (fin).

#### <a name="syntax"></a>Syntaxe
```js
rangeObject.select(selectionMode);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Facultatif. Le mode de sélection peut être « Select » (sélectionner), « Start » (début) ou « End » (fin). « Select » (sélectionner) est la valeur par défaut.|

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert HTML in to the beginning of the range.
    range.insertHtml('<strong>This is text inserted with range.insertHtml()</strong>', Word.InsertLocation.start);

    // Queue a command to select the HTML that was inserted.
    range.select();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Selected the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="support-details"></a>Informations de prise en charge
Utilisez l’[ensemble de conditions requises](../office-add-in-requirement-sets.md) dans les vérifications à l’exécution pour vous assurer que votre application est prise en charge par la version d’hôte de Word. Pour plus d’informations sur la configuration requise pour le serveur et l’application d’hôte Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).
