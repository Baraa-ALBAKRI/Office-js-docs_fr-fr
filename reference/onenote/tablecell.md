# <a name="tablecell-object-(javascript-api-for-onenote)"></a>Objet TableCell (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_  


Représente une cellule dans un tableau OneNote.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|cellIndex|int|Obtient l’index de la cellule dans la ligne correspondante. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-cellIndex)|
|id|chaîne|Obtient l’ID de la cellule. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-id)|
|rowIndex|int|Obtient l’index de la ligne de la cellule dans le tableau. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-rowIndex)|
|shadingColor|chaîne|Obtient et définit la couleur d’ombrage de la cellule|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-shadingColor)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type   |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Obtient la collection d’objets Paragraph dans TableCell. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-paragraphs)|
|parentRow|[TableRow](tablerow.md)|Obtient la ligne parent de la cellule. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-parentRow)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[appendHtml(html: string)](#appendhtmlhtml-string)|void|Ajoute le code HTML spécifié dans TableCell.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendHtml)|
|[appendImage(base64EncodedImage: string, width: double, height: double)](#appendimagebase64encodedimage-string-width-double-height-double)|[Image](image.md)|Ajoute l’image spécifiée à la cellule du tableau.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendImage)|
|[appendRichText(paragraphText: string)](#appendrichtextparagraphtext-string)|[RichText](richtext.md)|Ajoute le texte spécifié à la cellule du tableau.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendRichText)|
|[appendTable(rowCount: number, columnCount: number, values: string[][])](#appendtablerowcount-number-columncount-number-values-string)|[Table](table.md)|Ajoute un tableau avec le nombre spécifié de lignes et de colonnes dans une cellule du tableau.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendTable)|
|[clear()](#clear)|void|Efface le contenu de la cellule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-clear)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-load)|

## <a name="method-details"></a>Détails des méthodes


### <a name="appendhtml(html:-string)"></a>appendHtml(html: string)
Ajoute le code HTML spécifié dans TableCell.

#### <a name="syntax"></a>Syntaxe
```js
tableCellObject.appendHtml(html);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|Html|chaîne|Chaîne HTML à ajouter. Voir [HTML pris en charge](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) pour l’API JavaScript des compléments OneNote.|

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two and add "Hello".
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                cell.appendHtml("<p>Hello</p>");
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});


### appendImage(base64EncodedImage: string, width: double, height: double)
Adds the specified image to table cell.

#### Syntax
```js
tableCellObject.appendImage(base64EncodedImage, width, height);
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
Ajoute le texte spécifié à la cellule du tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableCellObject.appendRichText(paragraphText);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|paragraphText|string|Chaîne HTML à ajouter.|

#### <a name="returns"></a>Retourne
[RichText](richtext.md)

#### <a name="examples"></a>Exemples
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    var appendedRichText = null;
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two and add "Hello".
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                appendedRichText = cell.appendRichText("Hello");
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="appendtable(rowcount:-number,-columncount:-number,-values:-string[][])"></a>appendTable(rowCount: number, columnCount: number, values: string[][])
Ajoute un tableau avec le nombre spécifié de lignes et de colonnes dans une cellule du tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableCellObject.appendTable(rowCount, columnCount, values);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|rowCount|number|Obligatoire. Nombre de lignes dans le tableau.|
|columnCount|number|Obligatoire. Nombre de colonnes dans le tableau.|
|values|string[][]|Facultatif. Tableau 2D facultatif. Les cellules sont remplies si les chaînes correspondantes sont spécifiées dans le tableau.|

#### <a name="returns"></a>Retourne
[Table](table.md)

### <a name="clear()"></a>clear()
Efface le contenu de la cellule.

#### <a name="syntax"></a>Syntaxe
```js
tableCellObject.clear();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

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
**id, cellIndex, rowIndex**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                
                // Queue a command to load the table cell.
                ctx.load(cell);
                ctx.sync().then(function() {
                    console.log("Cell Id: " + cell.id);
                    console.log("Cell Index: " + cell.cellIndex);
                    console.log("Cell's Row Index: " + cell.rowIndex);
                });
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**parentTable, cells**
```js
ParentTable, ParentRow, Paragraphs
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                
                // Queue a command to load parentTable, parentRow and paragraphs of the table cell.
                ctx.load(cell, "parentTable, parentRow, paragraphs");
                
                ctx.sync().then(function() {
                    console.log("Parent Table Id: " + cell.parentTable.id);
                    console.log("Parent Row Id: " + cell.parentRow.id);
                    var paragraphs = cell.paragraphs;
                    
                    for (var i = 0; i < paragraphs.items.length; i++) {
                        console.log("Paragraph Id: " + paragraphs.items[i].id);
                    }
                });
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

