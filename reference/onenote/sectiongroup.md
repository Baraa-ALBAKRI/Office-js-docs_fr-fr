# <a name="sectiongroup-object-(javascript-api-for-onenote)"></a>Objet SectionGroup (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_   


Représente un groupe de sections OneNote. Les groupes de sections peuvent contenir des sections et des groupes de sections.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|clientUrl{|chaîne|URL du client du groupe de sections. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-clientUrl{)|
|id|string|Obtient l’ID du groupe de sections. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-id)|
|name|chaîne|Obtient le nom du groupe de sections. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-name)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type   |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|notebook|[Notebook](notebook.md)|Obtient le bloc-notes qui contient le groupe de sections. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-notebook)|
|parentSectionGroup|[SectionGroup](sectiongroup.md)|Obtient le groupe de sections qui contient le groupe de sections. Génère ItemNotFound si le groupe de sections est un enfant direct du bloc-notes. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-parentSectionGroup)|
|parentSectionGroupOrNull|[SectionGroup](sectiongroup.md)|Obtient le groupe de sections qui contient le groupe de sections. Renvoie la valeur Null si le groupe de sections est un enfant direct du bloc-notes. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-parentSectionGroupOrNull)|
|sectionGroups|[SectionGroupCollection](sectiongroupcollection.md)|Collection de groupes de sections dans le groupe de sections. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-sectionGroups)|
|Sections|[SectionCollection](sectioncollection.md)|Collection de sections dans le groupe de sections. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-sections)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[addSection(title: String)](#addsectiontitle-string)|[Section](section.md)|Ajoute une nouvelle section à la fin du groupe de sections.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-addSection)|
|[addSectionGroup(name: String)](#addsectiongroupname-string)|[SectionGroup](sectiongroup.md)|Ajoute un nouveau groupe de sections à la fin de cet objet sectionGroup.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-addSectionGroup)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-load)|

## <a name="method-details"></a>Détails des méthodes


### <a name="addsection(title:-string)"></a>addSection(title: String)
Ajoute une nouvelle section à la fin du groupe de sections.

#### <a name="syntax"></a>Syntaxe
```js
sectionGroupObject.addSection(title);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|title|String|Nom de la nouvelle section.|

#### <a name="returns"></a>Retourne
[Section](section.md)

#### <a name="examples"></a>Exemples
```js
OneNote.run(function (context) {

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;
    
    // Queue a command to load the section groups.
    // For best performance, request specific properties.
    sectionGroups.load("id");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            
            // Add a section to each section group.
            $.each(sectionGroups.items, function(index, sectionGroup) {
                sectionGroup.addSection("Agenda");
            });
            
            // Run the queued commands.
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


### <a name="addsectiongroup(name:-string)"></a>addSectionGroup(name: String)
Ajoute un nouveau groupe de sections à la fin de cet objet sectionGroup.

#### <a name="syntax"></a>Syntaxe
```js
sectionGroupObject.addSectionGroup(name);
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
    var sectionGroup;
    var nestedSectionGroup;

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();

    // Queue a command to add a new section group.
    var sectionGroups = notebook.sectionGroups;

    // Queue a command to load the new section group.
    sectionGroups.load();

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function(){
            sectionGroup = sectionGroups.items[0];
            sectionGroup.load();
            return context.sync();
        })
        .then(function(){
            nestedSectionGroup = sectionGroup.addSectionGroup("Sample nested section group");
            nestedSectionGroup.load();
            return context.sync();
        })
        .then(function() {
            console.log("New nested section group name is " + nestedSectionGroup.name);
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
        
    // Get the parent section group that contains the current section.
    var sectionGroup = context.application.getActiveSection().parentSectionGroup;
            
    // Queue a command to load the section group. 
    // For best performance, request specific properties.           
    sectionGroup.load("id,name");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Write the properties.
            console.log("Section group name: " + sectionGroup.name);
            console.log("Section group ID: " + sectionGroup.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**name et notebook**
```js
OneNote.run(function (context) {
        
    // Get the parent section group that contains the current section.
    var sectionGroup = context.application.getActiveSection().parentSectionGroup;
            
    // Queue a command to load the section group with the specified properties.           
    sectionGroup.load("name,notebook/name"); 
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Write the properties.
            console.log("Section group name: " + sectionGroup.name);
            console.log("Parent notebook name: " + sectionGroup.notebook.name);
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

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the section groups.
    // For best performance, request specific properties.
    sectionGroups.load("name");
    
    // Get the child section groups of the first section group in the notebook.
    var nestedSectionGroups = sectionGroups._GetItem(0).sectionGroups;
    
    // Queue a command to load the ID and name properties of the child section groups.
    nestedSectionGroups.load("id,name");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            
            // Write the properties for each child section group.
            $.each(nestedSectionGroups.items, function(index, sectionGroup) {
                console.log("Section group name: " + sectionGroup.name);  
                console.log("Section group ID: " + sectionGroup.id);  
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

    // Get the sections that are siblings of the current section.
    var sections = context.application.getActiveSection().parentSectionGroup.sections;

    // Queue a command to load the section groups.
    // For best performance, request specific properties.
    sections.load("id,name");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            
            // Write the properties for each section.
            $.each(sections.items, function(index, section) {
                console.log("Section name: " + section.name);  
                console.log("Section ID: " + section.id);  
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

