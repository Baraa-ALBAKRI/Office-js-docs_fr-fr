# <a name="contentcontrol-object-(javascript-api-for-word)"></a>Objet ContentControl (interface API JavaScript pour Word)

Représente un contrôle de contenu. Les contrôles de contenu sont des zones d’un document délimitées par des bordures et pouvant porter une étiquette qui servent à contenir certains types de contenu. Les contrôles de contenu individuels peuvent contenir des images, des tableaux ou des paragraphes de texte mis en forme. Actuellement, seuls les contrôles de contenu à texte enrichi sont pris en charge.

_S’applique à : Word 2016, Word pour iPad, Word pour Mac, Word Online_

## <a name="properties"></a>Propriétés
| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|cannotDelete|bool|Obtient ou définit une valeur qui indique si l’utilisateur peut supprimer le contrôle de contenu. Non compatible avec removeWhenEdited.|
|cannotEdit|bool|Obtient ou définit une valeur qui indique si l’utilisateur peut modifier le contenu du contrôle.|
|color|string|Obtient ou définit la couleur du contrôle de contenu. Celle-ci est définie au format « #RRGGBB » ou par le nom de la couleur.|
|placeholderText|string|Obtient ou définit le texte de l’espace réservé du contrôle de contenu. Ce texte apparaît de façon estompée lorsque le contrôle de contenu est vide. Cette propriété n’est actuellement pas prise en charge dans Word Online.|
|removeWhenEdited|bool|Obtient ou définit une valeur qui indique si le contrôle de contenu doit être supprimé après modification. Non compatible avec cannotDelete.|
|style|string|Obtient ou définit le style utilisé pour le contrôle de contenu. Il s’agit du nom du style pré-installé ou personnalisé.|
|tag|string|Obtient ou définit un indicateur pour identifier un contrôle de contenu. Le complément d’exemple [Silly stories](https://aka.ms/sillystorywordaddin) montre comment utiliser la propriété **tag**.|
|texte|string|Obtient le texte du contrôle de contenu. En lecture seule.|
|title|string|Obtient ou définit le titre d’un contrôle de contenu.|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type   |Description|
|:---------------|:--------|:----------|
|appearance|**ContentControlAppearance**|Obtient ou définit l’apparence du contrôle de contenu. La valeur peut être « boundingBox » (cadre englobant), « tags » (indicateurs) ou « hidden » (masqué).|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Obtient la collection d’objets de contrôle de contenu compris dans le contrôle de contenu. En lecture seule.|
|police|[Font](font.md)|Obtient le format de texte du contrôle de contenu. Utilisez cette propriété pour obtenir et définir le nom de la police, la taille, la couleur et d’autres propriétés. En lecture seule.|
|id|**uint**|Obtient un entier qui représente l’identificateur du contrôle de contenu. En lecture seule.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Obtient la collection d’objets inlinePicture du contrôle de contenu. La collection n’inclut pas d’images flottantes. En lecture seule.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Obtient la collection d’objets de paragraphe du contrôle de contenu. En lecture seule.|
|parentContentControl|[ContentControl](contentcontrol.md)|Obtient le contrôle de contenu qui contient le contrôle de contenu spécifié. Renvoie null s’il n’existe pas de contrôle de contenu parent. En lecture seule.|
|type|**ContentControlType**|Obtient le type du contrôle de contenu. Actuellement, seuls les contrôles de contenu à texte enrichi sont pris en charge. En lecture seule.|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Efface le contenu du contrôle de contenu. L’utilisateur peut effectuer l’opération d’annulation sur le contenu effacé.|
|[delete(keepContent: bool)](#deletekeepcontent-bool)|void|Supprime le contrôle de contenu et son contenu. Si keepContent est défini sur true, le contenu n’est pas supprimé.|
|[getHtml()](#gethtml)|chaîne|Obtient la représentation HTML de l’objet de contrôle de contenu.|
|[getOoxml()](#getooxml)|chaîne|Obtient la représentation Office Open XML (OOXML) de l’objet de contrôle de contenu.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Insère un saut à l’emplacement spécifié. Vous pouvez uniquement insérer un saut dans des objets qui sont contenus dans le corps de document principal, sauf s’il s’agit d’un saut de ligne, auquel cas il peut être inséré dans n’importe quel objet de corps. La valeur insertLocation peut être définie sur « Before » (avant), « After » (après), « Start » (début) ou « End » (fin).|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Insère un document dans le contrôle de contenu, à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Insère du code HTML dans le contrôle de contenu, à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Insère une image incluse dans le contrôle de contenu, à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin). |
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Insère du code OOXML ou un élément wordProcessingML dans le contrôle de contenu, à l’emplacement spécifié.  La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Insère un paragraphe à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant), « After » (après), « Start » (début) ou « End » (fin).|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Insère du texte dans le contrôle de contenu, à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestringssearchoptions)|[SearchResultCollection](searchresultcollection.md)|Effectue une recherche avec les options de recherche spécifiées dans l’étendue de l’objet de contrôle de contenu. Les résultats de la recherche sont un ensemble d’objets de plage.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Sélectionne le contrôle de contenu. Word fait défiler le document jusqu’à accéder à la sélection. Le mode de sélection peut être « Select » (sélectionner), « Start » (début) ou « End » (fin).|

## <a name="method-details"></a>Détails de méthodes

### <a name="clear()"></a>Effacer
Efface le contenu du contrôle de contenu. L’utilisateur peut effectuer l’opération d’annulation sur le contenu effacé.

#### <a name="syntax"></a>Syntaxe
```js
contentControlObject.clear();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the content controls collection.
    contentControls.load('text');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {
            
            // Queue a command to clear the contents of the first content control.
            contentControls.items[0].clear();
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });      
        }
            
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

### <a name="delete(keepcontent:-bool)"></a>delete(keepContent: bool)
Supprime le contrôle de contenu et son contenu. Si keepContent est défini sur true, le contenu n’est pas supprimé.

#### <a name="syntax"></a>Syntaxe
```js
contentControlObject.delete(keepContent);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|keepContent|bool|Obligatoire. Indique si le contenu doit être supprimé avec le contrôle de contenu. Si keepContent est défini sur true, le contenu n’est pas supprimé.|

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the content controls collection.
    contentControls.load('text');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {
            
            // Queue a command to delete the first content control. The
            // contents will remain in the document.
            contentControls.items[0].delete(true);
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });      
        }
            
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
Obtient la représentation HTML de l’objet de contrôle de contenu.

#### <a name="syntax"></a>Syntaxe
```js
contentControlObject.getHtml();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
string

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection that contains a specific tag.
    var contentControlsWithTag = context.document.contentControls.getByTag('Customer-Address');
    
    // Queue a command to load the tag property for all of content controls. 
    context.load(contentControlsWithTag, 'tag');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTag.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to get the HTML contents of the first content control.
            var html = contentControlsWithTag.items[0].getHtml();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Content control HTML: ' + html.value);
            });
        }
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
Obtient la représentation Office Open XML (OOXML) de l’objet de contrôle de contenu.

#### <a name="syntax"></a>Syntaxe
```js
contentControlObject.getOoxml();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
string

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to get the OOXML contents of the first content control.
            var ooxml = contentControls.items[0].getOoxml();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Content control OOXML: ' + ooxml.value);
            });
        }
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
Insère un saut à l’emplacement spécifié. Vous pouvez uniquement insérer un saut dans des objets qui sont contenus dans le corps de document principal, sauf s’il s’agit d’un saut de ligne, auquel cas il peut être inséré dans n’importe quel objet de corps. La valeur insertLocation peut être définie sur « Before » (avant), « After » (après), « Start » (début) ou « End » (fin).

#### <a name="syntax"></a>Syntaxe
```js
contentControlObject.insertBreak(breakType, insertLocation);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|breakType|BreakType|Obligatoire. Type de saut (breakType.md)|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être définie sur « Before » (avant), « After » (après), « Start » (début) ou « End » (fin).|

#### <a name="returns"></a>Retourne
void

#### <a name="additional-details"></a>Détails supplémentaires
À l’exception des sauts de ligne, vous ne pouvez pas insérer de saut dans les objets contenus dans les en-têtes, les pieds de page, les notes de bas de page, les notes de fin, les commentaires et les zones de texte.  

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a commmand to load the id property for all of content controls. 
    context.load(contentControls, 'id');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion. We now will have 
    // access to the content control collection.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to insert a page break after the first content control. 
            contentControls.items[0].insertBreak('page', "After");
            
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion. 
            return context.sync()
                .then(function () {
                    console.log('Inserted a page break after the first content control.');    
            });
        }
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
Insère un document dans le contrôle de contenu, à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).

#### <a name="syntax"></a>Syntaxe
```js
contentControlObject.insertFileFromBase64(base64File, insertLocation);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|base64File|string|Obligatoire. Contenu de fichier encodé au format Base64 à insérer.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|

#### <a name="returns"></a>Retourne
[Range](range.md)

### <a name="inserthtml(html:-string,-insertlocation:-insertlocation)"></a>insertHtml(html: string, insertLocation: InsertLocation)
Insère du code HTML dans le contrôle de contenu, à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).

#### <a name="syntax"></a>Syntaxe
```js
contentControlObject.insertHtml(html, insertLocation);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|Html|string|Obligatoire. Code HTML à insérer dans le contrôle de contenu.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to put HTML into the contents of the first content control.
            contentControls.items[0].insertHtml('<strong>HTML content inserted into the content control.</strong>', 'Start');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted HTML in the first content control.');
            });
        }
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
Insère une image incluse dans le contrôle de contenu, à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).

#### <a name="syntax"></a>Syntaxe
contentControlObject.insertInlinePictureFromBase64(image, insertLocation);

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Obligatoire. Image encodée au format Base64 à insérer dans le contrôle de contenu.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|

#### <a name="returns"></a>Retourne
[InlinePicture](inlinepicture.md)

#### <a name="known-issues"></a>Problèmes connus
Dans Word Online, seule la valeur « Replace » (remplacer) est prise en charge pour le paramètre _insertLocation_. Si vous utilisez les valeurs « Start » (début) ou « End » (fin), l’opération échoue.

### <a name="insertooxml(ooxml:-string,-insertlocation:-insertlocation)"></a>insertOoxml(ooxml: string, insertLocation: InsertLocation)
Insère du contenu OOXML ou wordProcessingML dans le contrôle de contenu, à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).

#### <a name="syntax"></a>Syntaxe
```js
contentControlObject.insertOoxml(ooxml, insertLocation);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|ooxml|string|Obligatoire. Contenu OOXML ou wordProcessingML à insérer dans le contrôle de contenu.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="known-issues"></a>Problèmes connus
Cette méthode engendre une longue latence dans Word Online, ce qui peut affecter l’expérience des utilisateurs de votre complément. Nous vous recommandons d’utiliser cette méthode uniquement lorsqu’aucune solution n’est disponible. 

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to put OOXML into the contents of the first content control.
            contentControls.items[0].insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", "End");
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted OOXML in the first content control.');
            });
        }
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
Insère un paragraphe à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant), « After » (après), « Start » (début) ou « End » (fin).

#### <a name="syntax"></a>Syntaxe
```js
contentControlObject.insertParagraph(paragraphText, insertLocation);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|paragraphText|string|Obligatoire. Texte de paragraphe à insérer.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être définie sur « Before » (avant), « After » (après), « Start » (début) ou « End » (fin).|

#### <a name="returns"></a>Retourne
[Paragraph](paragraph.md)

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to insert a paragraph after the first content control. 
            contentControls.items[0].insertParagraph('Text of the inserted paragraph.', 'After');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted a paragraph after the first content control.');
            });
        }
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
Insère du texte dans le contrôle de contenu, à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).

#### <a name="syntax"></a>Syntaxe
```js
contentControlObject.insertText(text, insertLocation);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|text|string|Obligatoire. Texte à insérer dans le contrôle de contenu.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="known-issues"></a>Problèmes connus
Dans Word Online, seule la valeur « Replace » (remplacer) est prise en charge pour le paramètre _insertLocation_. Si vous utilisez les valeurs « Start » (début) ou « End » (fin), l’opération échoue.

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to replace text in the first content control. 
            contentControls.items[0].insertText('Replaced text in the first content control.', 'Replace');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Replaced text in the first content control.');
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

Le complément d’exemple [Silly stories](https://aka.ms/sillystorywordaddin) montre comment utiliser la méthode **insertText**.

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
    
    // Create a proxy range object for the current selection.
    var range = context.document.getSelection();
    
    // Queue a commmand to create the content control.
    var myContentControl = range.insertContentControl();
    myContentControl.tag = 'Customer-Address';
    myContentControl.title = ' has t';
    myContentControl.style = 'Heading 2';
    myContentControl.insertText('One Microsoft Way, Redmond, WA 98052', 'replace');
    myContentControl.cannotEdit = true;
    
    // Queue a command to load the id property for the content control you created.
    context.load(myContentControl, 'id');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Created content control with id: ' + myContentControl.id);
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
Effectue une recherche avec les options de recherche spécifiées dans l’étendue de l’objet de contrôle de contenu. Les résultats de la recherche sont un ensemble d’objets de plage.

#### <a name="syntax"></a>Syntaxe
```js
contentControlObject.search(searchText, searchOptions);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|searchText|string|Obligatoire. Texte de recherche.|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|Facultatif. Options de la recherche.|

#### <a name="returns"></a>Retourne
[SearchResultCollection](searchresultcollection.md)

### <a name="select(selectionmode:-selectionmode)"></a>select(selectionMode: SelectionMode)
Sélectionne le contrôle de contenu. Word fait défiler le document jusqu’à accéder à la sélection. Le mode de sélection peut être « Select » (sélectionner), « Start » (début) ou « End » (fin).

#### <a name="syntax"></a>Syntaxe
```js
contentControlObject.select(selectionMode);
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
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to select the first content control.
            contentControls.items[0].select();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Selected the first content control.');
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="property-access-examples"></a>Exemples d’accès aux propriétés

### <a name="load-all-of-the-content-control-properties"></a>Charger toutes les propriétés du contrôle de contenu
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to load the properties on the first content control. 
            contentControls.items[0].load(  'appearance,' +
                                            'cannotDelete,' +
                                            'cannotEdit,' +
                                            'id,' +
                                            'placeHolderText,' +
                                            'removeWhenEdited,' +
                                            'title,' +
                                            'text,' +
                                            'type,' +
                                            'style,' +
                                            'tag,' +
                                            'font/size,' +
                                            'font/name,' +
                                            'font/color');             
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Property values of the first content control:' + 
                        '   ----- appearance: ' + contentControls.items[0].appearance + 
                        '   ----- cannotDelete: ' + contentControls.items[0].cannotDelete +
                        '   ----- cannotEdit: ' + contentControls.items[0].cannotEdit +
                        '   ----- color: ' + contentControls.items[0].color +
                        '   ----- id: ' + contentControls.items[0].id +
                        '   ----- placeHolderText: ' + contentControls.items[0].placeholderText +
                        '   ----- removeWhenEdited: ' + contentControls.items[0].removeWhenEdited +
                        '   ----- title: ' + contentControls.items[0].title +
                        '   ----- text: ' + contentControls.items[0].text +
                        '   ----- type: ' + contentControls.items[0].type +
                        '   ----- style: ' + contentControls.items[0].style +
                        '   ----- tag: ' + contentControls.items[0].tag +
                        '   ----- font size: ' + contentControls.items[0].font.size +
                        '   ----- font name: ' + contentControls.items[0].font.name +
                        '   ----- font color: ' + contentControls.items[0].font.color);
            });
        }
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
