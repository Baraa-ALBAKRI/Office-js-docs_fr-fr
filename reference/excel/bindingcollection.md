# <a name="bindingcollection-object-javascript-api-for-excel"></a>Objet BindingCollection (API JavaScript pour Excel)

Représente la collection de l’ensemble des objets de liaison qui font partie du classeur.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|count|int|Renvoie le nombre de liaisons de la collection. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|éléments|[Binding[]](binding.md)|Collection d’objets de liaison. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[add(range: plage ou chaîne, bindingType: chaîne, id: chaîne)](#addrange-range-or-string-bindingtype-string-id-string)|[Binding](binding.md)|Ajouter une nouvelle liaison à une plage spécifique.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[addFromNamedItem(name: chaîne, bindingType: chaîne, id: chaîne)](#addfromnameditemname-string-bindingtype-string-id-string)|[Binding](binding.md)|Ajouter une nouvelle liaison basée sur un élément nommé dans le classeur.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[addFromSelection(bindingType: chaîne, id: chaîne)](#addfromselectionbindingtype-string-id-string)|[Binding](binding.md)|Ajouter une nouvelle liaison basée sur la sélection en cours.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Obtient le nombre de liaisons de la collection.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(id: string)](#getitemid-string)|[Binding](binding.md)|Obtient un objet de liaison par ID.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Binding](binding.md)|Obtient un objet de liaison en fonction de sa position dans le tableau d’éléments.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(id: string)](#getitemornullobjectid-string)|[Binding](binding.md)|Obtient un objet de liaison par ID. Si l’objet de liaison n’existe pas, renvoie un objet null.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="addrange-range-or-string-bindingtype-string-id-string"></a>add(range: plage ou chaîne, bindingType: chaîne, id: chaîne)
Ajouter une nouvelle liaison à une plage spécifique.

#### <a name="syntax"></a>Syntaxe
```js
bindingCollectionObject.add(range, bindingType, id);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|plage|range ou string|Plage à laquelle lier la liaison. Peut être un objet de plage Excel ou une chaîne. Si c’est une chaîne, elle doit contenir l’adresse complète, y compris le nom de la feuille.|
|bindingType|string|Type de liaison.  Les valeurs possibles sont les suivantes : Range, Table, Text|
|id|chaîne|Nom de la liaison.|

#### <a name="returns"></a>Retourne
[Binding](binding.md)

### <a name="addfromnameditemname-string-bindingtype-string-id-string"></a>addFromNamedItem(name: chaîne, bindingType: chaîne, id: chaîne)
Ajouter une nouvelle liaison basée sur un élément nommé dans le classeur.

#### <a name="syntax"></a>Syntaxe
```js
bindingCollectionObject.addFromNamedItem(name, bindingType, id);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|name|chaîne|Nom à partir duquel créer la liaison.|
|bindingType|string|Type de liaison.  Les valeurs possibles sont les suivantes : Range, Table, Text|
|id|chaîne|Nom de la liaison.|

#### <a name="returns"></a>Retourne
[Binding](binding.md)

### <a name="addfromselectionbindingtype-string-id-string"></a>addFromSelection(bindingType: chaîne, id: chaîne)
Ajouter une nouvelle liaison basée sur la sélection en cours.

#### <a name="syntax"></a>Syntaxe
```js
bindingCollectionObject.addFromSelection(bindingType, id);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|bindingType|string|Type de liaison.  Les valeurs possibles sont les suivantes : Range, Table, Text|
|id|chaîne|Nom de la liaison.|

#### <a name="returns"></a>Retourne
[Binding](binding.md)

### <a name="getcount"></a>getCount()
Obtient le nombre de liaisons de la collection.

#### <a name="syntax"></a>Syntaxe
```js
bindingCollectionObject.getCount();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Renvoie
int

### <a name="getitemid-string"></a>getItem(id: string)
Obtient un objet de liaison par ID.

#### <a name="syntax"></a>Syntaxe
```js
bindingCollectionObject.getItem(id);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|id|string|ID de l’objet de liaison à récupérer.|

#### <a name="returns"></a>Retourne
[Binding](binding.md)

#### <a name="examples"></a>Exemples

Créez une liaison de table pour contrôler les modifications apportées aux données de la table. Lorsque les données sont modifiées, la couleur d’arrière-plan du tableau devient orange.

```js
function addEventHandler() {
    //Create Table1
Excel.run(function (ctx) { 
    ctx.workbook.tables.add("Sheet1!A1:C4", true);
    return ctx.sync().then(function() {
             console.log("My Diet Data Inserted!");
    })
    .catch(function (error) {
             console.log(JSON.stringify(error));
    });
});
    //Create a new table binding for Table1
Office.context.document.bindings.addFromNamedItemAsync("Table1", Office.CoercionType.Table, { id: "myBinding" }, function (asyncResult) {
    if (asyncResult.status == "failed") {
        console.log("Action failed with error: " + asyncResult.error.message);
    }
    else {
        // If succeeded, then add event handler to the table binding.
        Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
    }
});
}
    
// when data in the table is changed, this event will be triggered.
function onBindingDataChanged(eventArgs) {
Excel.run(function (ctx) { 
    // highlight the table in orange to indicate data has been changed.
    ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.color = "Orange";
    return ctx.sync().then(function() {
            console.log("The value in this table got changed!");
    })
    .catch(function (error) {
            console.log(JSON.stringify(error));
    });
});
}

```



#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.bindings.count - 1;
    var binding = ctx.workbook.bindings.getItemAt(lastPosition);
    binding.load('type')
    return ctx.sync().then(function() {
            console.log(binding.type); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitematindex-number"></a>getItemAt(index: number)
Obtient un objet de liaison en fonction de sa position dans le tableau d’éléments.

#### <a name="syntax"></a>Syntaxe
```js
bindingCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[Binding](binding.md)

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.bindings.count - 1;
    var binding = ctx.workbook.bindings.getItemAt(lastPosition);
    binding.load('type')
    return ctx.sync().then(function() {
            console.log(binding.type); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitemornullobjectid-string"></a>getItemOrNullObject(id: string)
Obtient un objet de liaison par ID. Si l’objet de liaison n’existe pas, renvoie un objet null.

#### <a name="syntax"></a>Syntaxe
```js
bindingCollectionObject.getItemOrNullObject(id);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|id|string|ID de l’objet de liaison à récupérer.|

#### <a name="returns"></a>Retourne
[Binding](binding.md)
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

```js
Excel.run(function (ctx) { 
    var bindings = ctx.workbook.bindings;
    bindings.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < bindings.items.length; i++)
        {
            console.log(bindings.items[i].id);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
Obtenir le nombre de liaisons

```js
Excel.run(function (ctx) { 
    var bindings = ctx.workbook.bindings;
    bindings.load('count');
    return ctx.sync().then(function() {
        console.log("Bindings: Count= " + bindings.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
