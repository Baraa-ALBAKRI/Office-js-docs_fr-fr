# <a name="sectioncollection-object-(javascript-api-for-onenote)"></a>Objet SectionCollection (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_  


Représente une collection de sections.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|count|int|Renvoie le nombre de sections de la collection. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-count)|
|items|[Section[]](section.md)|Collection d’objets de section. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-items)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[getByName(name: string)](#getbynamename-string)|[SectionCollection](sectioncollection.md)|Obtient la collection de sections portant le nom spécifié.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-getByName)|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[Section](section.md)|Obtient une section en fonction de son ID ou de son index dans la collection. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[Section](section.md)|Obtient une section en fonction de sa position dans la collection.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-load)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getbyname(name:-string)"></a>getByName(name: string)
Obtient la collection de sections portant le nom spécifié.

#### <a name="syntax"></a>Syntaxe
```js
sectionCollectionObject.getByName(name);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|name|string|Nom de la section.|

#### <a name="returns"></a>Retourne
[SectionCollection](sectioncollection.md)

#### <a name="examples"></a>Exemples
```js
OneNote.run(function (context) {

    // Get the sections in the current notebook.
    var sections = context.application.getActiveNotebook().sections;

    // Queue a command to load the sections. 
    // For best performance, request specific properties.
    sections.load("id"); 
    
    // Get the sections with the specified name.
    var groceriesSections = sections.getByName("Groceries");
    
    // Queue a command to load the sections with the specified name.
    groceriesSections.load("id,name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index.
            if (groceriesSections.items.length > 0) {
                console.log("Section name: " + groceriesSections.items[0].name);
                console.log("Section ID: " + groceriesSections.items[0].id);
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
Obtient une section en fonction de son ID ou de son index dans la collection. En lecture seule.

#### <a name="syntax"></a>Syntaxe
```js
sectionCollectionObject.getItem(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number or string|ID ou emplacement d’index de la section de la collection.|

#### <a name="returns"></a>Retourne
[Section](section.md)

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
Obtient une section en fonction de sa position dans la collection.

#### <a name="syntax"></a>Syntaxe
```js
sectionCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[Section](section.md)

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

    // Get the sections in the current notebook.
    var sections = context.application.getActiveNotebook().sections;

    // Queue a command to load the sections. 
    // For best performance, request specific properties.
    sections.load("name"); 

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Iterate through the collection or access items individually by index, for example: sections.items[0]
            $.each(sections.items, function(index, section) {
                if (section.name === "Homework") {
                    section.addPage("Biology");
                    section.addPage("Spanish");
                    section.addPage("Computer Science");
                }
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

