# <a name="notebook-object-(javascript-api-for-onenote)"></a>Objet Notebook (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_   


Représente un bloc-notes OneNote. Les blocs-notes contiennent des groupes de sections et des sections.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|clientUrl|chaîne|URL du client du bloc-notes. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-clientUrl)|
|id|chaîne|Obtient l’ID du bloc-notes. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-id)|
|name|chaîne|Obtient le nom du bloc-notes. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-name)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type   |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|sectionGroups|[SectionGroupCollection](sectiongroupcollection.md)|Obtient les groupes de sections dans le bloc-notes. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-sectionGroups)|
|Sections|[SectionCollection](sectioncollection.md)|Sections du bloc-notes. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-sections)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[addSection(name: String)](#addsectionname-string)|[Section](section.md)|Ajoute une nouvelle section à la fin du bloc-notes.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-addSection)|
|[addSectionGroup(name: String)](#addsectiongroupname-string)|[SectionGroup](sectiongroup.md)|Ajoute un nouveau groupe de sections à la fin du bloc-notes.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-addSectionGroup)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-load)|

## <a name="method-details"></a>Détails des méthodes


### <a name="addsection(name:-string)"></a>addSection(name: String)
Ajoute une nouvelle section à la fin du bloc-notes.

#### <a name="syntax"></a>Syntaxe
```js
notebookObject.addSection(name);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|name|String|Nom de la nouvelle section.|

#### <a name="returns"></a>Retourne
[Section](section.md)

#### <a name="examples"></a>Exemples
```js          
OneNote.run(function (context) {

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();

    // Queue a command to add a new section. 
    var section = notebook.addSection("Sample section");
    
    // Queue a command to load the new section. This example reads the name property later.
    section.load("name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log("New section name is " + section.name);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```


### <a name="addsectiongroup(name:-string)"></a>addSectionGroup(name: String)
Ajoute un nouveau groupe de sections à la fin du bloc-notes.

#### <a name="syntax"></a>Syntaxe
```js
notebookObject.addSectionGroup(name);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|name|String|Nom de la nouvelle section.|

#### <a name="returns"></a>Retourne
[SectionGroup](sectiongroup.md)

#### <a name="examples"></a>Exemples
```js          
OneNote.run(function (context) {

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();

    // Queue a command to add a new section group.
    var sectionGroup = notebook.addSectionGroup("Sample section group");

    // Queue a command to load the new section group.
    sectionGroup.load();

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log("New section group name is " + sectionGroup.name);
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
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés
**id**
```js
OneNote.run(function (context) {
        
    // Get the current notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
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

**name**
```js
OneNote.run(function (context) {
        
    // Get the current notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("Notebook name: " + notebook.name);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**sectionGroups**
```js          
OneNote.run(function (context) {

    // Get the section groups in the notebook. 
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the sectionGroups. 
    sectionGroups.load("name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            $.each(sectionGroups.items, function(index, sectionGroup) {
                console.log("Section group name: " + sectionGroup.name);
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

**sections**
```js
OneNote.run(function (context) {

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();
    
    // Queue a command to get immediate child sections of the notebook. 
    var childSections = notebook.sections;

    // Queue a command to load the childSections. 
    context.load(childSections);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            $.each(childSections.items, function(index, childSection) {
                console.log("Immediate child section name: " + childSection.name);
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

