# <a name="bindingcollection-object-(javascript-api-for-excel)"></a>Objet BindingCollection (interface API JavaScript pour Excel)

Représente la collection de tous les objets de liaison qui font partie du classeur.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|count|int|Renvoie le nombre de liaisons de la collection. En lecture seule.|
|items|[Binding[]](binding.md)|Collection d’objets de liaison. En lecture seule.|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[getItem(id: string)](#getitemid-string)|[Binding](binding.md)|Obtient un objet de liaison par ID.|
|[getItemAt(index: number)](#getitematindex-number)|[Binding](binding.md)|Obtient un objet de liaison en fonction de sa position dans le tableau d’éléments.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## <a name="method-details"></a>Détails des méthodes


### <a name="getitem(id:-string)"></a>getItem(id: string)
Obtient un objet de liaison par ID.

#### <a name="syntax"></a>Syntaxe
```js
bindingCollectionObject.getItem(id);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|id|string|ID de l’objet de liaison à récupérer.|

#### <a name="returns"></a>Retourne
[Binding](binding.md)

#### <a name="examples"></a>Exemples

Créez une liaison de table pour contrôler les modifications apportées aux données de la table. Lorsque les données sont modifiées, la couleur d’arrière-plan du tableau devient orange.

```js
(function () {
    // Create myTable
    Excel.run(function (ctx) {
        var table = ctx.workbook.tables.add("Sheet1!A1:C4", true);
        table.name = "myTable";
        return ctx.sync().then(function () {
            console.log("MyTable is Created!");

            //Create a new table binding for myTable
            Office.context.document.bindings.addFromNamedItemAsync("myTable", Office.CoercionType.Table, { id: "myBinding" }, function (asyncResult) {
                if (asyncResult.status == "failed") {
                    console.log("Action failed with error: " + asyncResult.error.message);
                }
                else {
                    // If successful, add the event handler to the table binding.
                    Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
                }
            });
        })
        .catch(function (error) {
            console.log(JSON.stringify(error));
        });
    });
    
    // When data in the table is changed, this event is triggered.
    function onBindingDataChanged(eventArgs) {
        Excel.run(function (ctx) {
            // Highlight the table in orange to indicate data changed.
            var fill = ctx.workbook.tables.getItem("myTable").getDataBodyRange().format.fill;
            fill.load("color");
            return ctx.sync().then(function () {
                if (fill.color != "Orange") {
                    ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.color = "Orange";
 
                    console.log("The value in this table got changed!");
                }
                else
                    
            })
                .then(ctx.sync)
            .catch(function (error) {
                console.log(JSON.stringify(error));
            });
        });
    } 
})();
 


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


### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
Obtient un objet de liaison en fonction de sa position dans le tableau d’éléments.

#### <a name="syntax"></a>Syntaxe
```js
bindingCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
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


### <a name="load(param:-object)"></a>load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### <a name="syntax"></a>Syntaxe
```js
object.load(param);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, accepte un objet [loadOption](loadoption.md).|

#### <a name="returns"></a>Renvoie
void
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

```js
Excel.run(function (ctx) { 
    var bindings = ctx.workbook.bindings;
    bindings.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < bindings.items.length; i++)
        {
            console.log(bindings.items[i].id);
            console.log(bindings.items[i].index);
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
