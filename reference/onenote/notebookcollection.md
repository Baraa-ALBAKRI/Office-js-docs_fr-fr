# <a name="notebookcollection-object-(javascript-api-for-onenote)"></a>Objet NotebookCollection (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_  


Représente une collection de blocs-notes.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|count|int|Renvoie le nombre de blocs-notes de la collection. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-count)|
|items|[Notebook[]](notebook.md)|Collection d’objets de bloc-notes. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-items)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[getByName(name: string)](#getbynamename-string)|[NotebookCollection](notebookcollection.md)|Obtient la collection de blocs-notes portant le nom spécifié qui sont ouverts dans l’instance de l’application.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-getByName)|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[Notebook](notebook.md)|Obtient un bloc-notes en fonction de son ID ou de son index dans la collection. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[Notebook](notebook.md)|Obtient un bloc-notes en fonction de sa position dans la collection.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-load)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getbyname(name:-string)"></a>getByName(name: string)
Obtient la collection de blocs-notes portant le nom spécifié qui sont ouverts dans l’instance de l’application.

#### <a name="syntax"></a>Syntaxe
```js
notebookCollectionObject.getByName(name);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|name|string|Nom du bloc-notes.|

#### <a name="returns"></a>Retourne
[NotebookCollection](notebookcollection.md)

#### <a name="examples"></a>Exemples
```js
OneNote.run(function (context) {

    // Get the notebooks that are open in the application instance and have the specified name.
    var notebooks = context.application.notebooks.getByName("Homework");

    // Queue a command to load the notebooks. 
    // For best performance, request specific properties.           
    notebooks.load("id,name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index, for example: notebooks.items[0]
            if (notebooks.items.length > 0) {
                console.log("Notebook name: " + notebooks.items[0].name);
                console.log("Notebook ID: " + notebooks.items[0].id);
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

### <a name="getitem(index:-number-or-string)"></a>getItem(index: number or string)
Obtient un bloc-notes en fonction de son ID ou de son index dans la collection. En lecture seule.

#### <a name="syntax"></a>Syntaxe
```js
notebookCollectionObject.getItem(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number or string|ID ou emplacement d’index du bloc-notes dans la collection.|

#### <a name="returns"></a>Retourne
[Notebook](notebook.md)

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
Obtient un bloc-notes en fonction de sa position dans la collection.

#### <a name="syntax"></a>Syntaxe
```js
notebookCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[Notebook](notebook.md)

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

**items**
```js
OneNote.run(function (context) {

    // Get the notebooks that are open in the application instance and have the specified name.
    var notebooks = context.application.notebooks.getByName("Homework");

    // Queue a command to load the notebooks. 
    // For best performance, request specific properties.           
    notebooks.load("id");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index, for example: notebooks.items[0]
            $.each(notebooks.items, function(index, notebook) {
                notebook.addSection("Biology");
                notebook.addSection("Spanish");
                notebook.addSection("Computer Science");
            });
            
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

