# <a name="application-object-javascript-api-for-onenote"></a>Objet Application (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_


Représente l’objet de niveau supérieur qui contient tous les objets OneNote globalement adressables tels que les blocs-notes, le bloc-notes actif et la section active.

## <a name="properties"></a>Propriétés

Aucun

## <a name="relationships"></a>Relations
| Relation | Type   |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|Blocs-notes|[NotebookCollection](notebookcollection.md)|Obtient la collection de blocs-notes ouverts dans l’instance de l’application OneNote. Dans OneNote Online, un seul bloc-notes est ouvert à la fois dans l’instance de l’application. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-notebooks)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[getActiveNotebook()](#getactivenotebook)|[Notebook](notebook.md)|Obtient le bloc-notes actif s’il existe. Si aucun bloc-notes n’est actif, génère ItemNotFound.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveNotebook)|
|[getActiveNotebookOrNull()](#getactivenotebookornull)|[Notebook](notebook.md)|Obtient le bloc-notes actif s’il existe. Si aucun bloc-notes n’est actif, renvoie la valeur Null.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveNotebookOrNull)|
|[getActiveOutline()](#getactiveoutline)|[Plan](outline.md)|Obtient le plan actif s’il existe. Si aucun plan n’est actif, génère ItemNotFound.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveOutline)|
|[getActiveOutlineOrNull()](#getactiveoutlineornull)|[Plan](outline.md)|Obtient le plan actif s’il existe. Sinon, renvoie la valeur Null.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveOutlineOrNull)|
|[getActivePage()](#getactivepage)|[Page](page.md)|Obtient la page active si elle existe. Si aucune page n’est active, génère ItemNotFound.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActivePage)|
|[getActivePageOrNull()](#getactivepageornull)|[Page](page.md)|Obtient la page active si elle existe. Si aucune page n’est active, renvoie la valeur Null.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActivePageOrNull)|
|[getActiveSection()](#getactivesection)|[Section](section.md)|Obtient la section active si elle existe. Si aucune section n’est active, génère ItemNotFound.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveSection)|
|[getActiveSectionOrNull()](#getactivesectionornull)|[Section](section.md)|Obtient la section active si elle existe. Si aucune section n’est active, renvoie la valeur Null.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveSectionOrNull)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-load)|
|[navigateToPage(page : Page)](#navigatetopagepage-page)|void|Ouvre la page spécifiée dans l’instance de l’application.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-navigateToPage)|
|[navigateToPageWithClientUrl(url: string)](#navigatetopagewithclienturlurl-string)|[Page](page.md)|Obtient la page spécifiée et ouvre celle-ci dans l’instance de l’application.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-navigateToPageWithClientUrl)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getactivenotebook"></a>getActiveNotebook()
Obtient le bloc-notes actif s’il existe. Si aucun bloc-notes n’est actif, génère ItemNotFound.

#### <a name="syntax"></a>Syntaxe
```js
applicationObject.getActiveNotebook();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Notebook](notebook.md)

#### <a name="examples"></a>Exemples
```js
OneNote.run(function (context) {
        
    // Get the active notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Notebook name: " + notebook.name);
            console.log("Notebook ID: " + notebook.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="getactivenotebookornull"></a>getActiveNotebookOrNull()
Obtient le bloc-notes actif s’il existe. Si aucun bloc-notes n’est actif, renvoie la valeur Null.

#### <a name="syntax"></a>Syntaxe
```js
applicationObject.getActiveNotebookOrNull();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Notebook](notebook.md)

#### <a name="examples"></a>Exemples
```js
OneNote.run(function (context) {

    // Get the active notebook.
    var notebook = context.application.getActiveNotebookOrNull();

    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id,name');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // check if active notebook is set.
            if (!notebook.isNull) {
                console.log("Notebook name: " + notebook.name);
                console.log("Notebook ID: " + notebook.id);
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


### <a name="getactiveoutline"></a>getActiveOutline()
Obtient le plan actif s’il existe. Si aucun plan n’est actif, génère ItemNotFound.

#### <a name="syntax"></a>Syntaxe
```js
applicationObject.getActiveOutline();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Outline](outline.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // get active outline.
    var outline = context.application.getActiveOutline();

    // Queue a command to load the id of the outline.         
    outline.load('id');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Show some properties.
            console.log("outline id: " + outline.id);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="getactiveoutlineornull"></a>getActiveOutlineOrNull()
Obtient le plan actif s’il existe. Sinon, renvoie la valeur Null.

#### <a name="syntax"></a>Syntaxe
```js
applicationObject.getActiveOutlineOrNull();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Outline](outline.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // get active outline.
    var outline = context.application.getActiveOutlineOrNull();

    // Queue a command to load the id of the outline.         
    outline.load('id');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            if (!outline.isNull) {
                console.log("outline id: " + outline.id);
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


### <a name="getactivepage"></a>getActivePage()
Obtient la page active si elle existe. Si aucune page n’est active, génère ItemNotFound.

#### <a name="syntax"></a>Syntaxe
```js
applicationObject.getActivePage();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Page](page.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {
        
    // Get the active page.
    var page = context.application.getActivePage();
            
    // Queue a command to load the page. 
    // For best performance, request specific properties.           
    page.load('id,title');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Page title: " + page.title);
            console.log("Page ID: " + page.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="getactivepageornull"></a>getActivePageOrNull()
Obtient la page active si elle existe. Si aucune page n’est active, renvoie la valeur Null.

#### <a name="syntax"></a>Syntaxe
```js
applicationObject.getActivePageOrNull();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Page](page.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // Get the active page.
    var page = context.application.getActivePageOrNull();

    // Queue a command to load the page. 
    // For best performance, request specific properties.           
    page.load('id,title');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            if (!page.isNull) {
                // Show some properties.
                console.log("Page title: " + page.title);
                console.log("Page ID: " + page.id);
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


### <a name="getactivesection"></a>getActiveSection()
Obtient la section active si elle existe. Si aucune section n’est active, génère ItemNotFound.

#### <a name="syntax"></a>Syntaxe
```js
applicationObject.getActiveSection();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Section](section.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {
        
    // Get the active section.
    var section = context.application.getActiveSection();
            
    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Section name: " + section.name);
            console.log("Section ID: " + section.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="getactivesectionornull"></a>getActiveSectionOrNull()
Obtient la section active si elle existe. Si aucune section n’est active, renvoie la valeur Null.

#### <a name="syntax"></a>Syntaxe
```js
applicationObject.getActiveSectionOrNull();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Section](section.md)

#### <a name="examples"></a>Exemples
```js
OneNote.run(function (context) {

    // Get the active section.
    var section = context.application.getActiveSectionOrNull();

    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load('id,name');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if (!section.isNull) {
                // Show some properties.
                console.log("Section name: " + section.name);
                console.log("Section ID: " + section.id);
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


### <a name="loadparam-object"></a>load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### <a name="syntax"></a>Syntaxe
```js
object.load(param);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### <a name="returns"></a>Renvoie
void

### <a name="navigatetopagepage-page"></a>navigateToPage(page: Page)
Ouvre la page spécifiée dans l’instance de l’application.

#### <a name="syntax"></a>Syntaxe
```js
applicationObject.navigateToPage(page);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|page|Page|Page à ouvrir.|

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js        
OneNote.run(function (context) {
        
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
            
    // Queue a command to load the pages. 
    // For best performance, request specific properties.           
    pages.load('id');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // This example loads the first page in the section.
            var page = pages.items[0];
                        
            // Open the page in the application.                    
            context.application.navigateToPage(page);
                    
            // Run the queued command.
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


### <a name="navigatetopagewithclienturlurl-string"></a>navigateToPageWithClientUrl(url: string)
Obtient la page spécifiée et ouvre celle-ci dans l’instance de l’application.

#### <a name="syntax"></a>Syntaxe
```js
applicationObject.navigateToPageWithClientUrl(url);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|url|string|URL du client de la page à ouvrir.|

#### <a name="returns"></a>Retourne
[Page](page.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;

    // Queue a command to load the pages. 
    // For best performance, request specific properties.           
    pages.load('clientUrl');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // This example loads the first page in the section.
            var page = pages.items[0];

            // Open the page in the application.                    
            context.application.navigateToPageWithClientUrl(page.clientUrl);

            // Run the queued command.
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
