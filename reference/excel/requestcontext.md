# <a name="requestcontext-object-(javascript-api-for-excel)"></a>Objet RequestContext (interface API JavaScript pour Excel)

l’objet de contexte de demande facilite les demandes auprès de l’application Excel. L’exécution du complément Office et de l’application Excel faisant appel à deux processus différents, il est nécessaire de fournir le contexte des demandes pour accéder à Excel et aux objets associés, tels que les feuilles de calcul, les tableaux, etc. à partir du complément. 

## <a name="properties"></a>Propriétés
Aucun

## <a name="methods"></a>Méthodes

| Méthode         | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[load(object: object, option: object)](#loadobject-object-option-object)  |void     |Insère l’objet de proxy créé dans le calque JavaScript avec les propriétés et les options spécifiées dans le paramètre.|

## <a name="api-specification"></a>Spécification d’API

### <a name="load(object:-object,-option:-object)"></a>load(object: object, option: object)
Insère l’objet de proxy créé dans le calque JavaScript avec les propriétés et les options spécifiées dans le paramètre.

#### <a name="syntax"></a>Syntaxe
```js
requestContextObject.load(object, loadOption);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:----------------|:--------|:----------|
|object|object|Facultatif. Indiquez le nom de l’objet à charger.|
|Option|[loadOption](loadoption.md)|Facultatif. Spécifiez les options de chargement (select, expand, skip ou top). Pour plus d’informations, voir l’objet loadOption.|

#### <a name="returns"></a>Retourne
void

##### <a name="examples"></a>Exemples

L’exemple suivant charge les valeurs de propriété d’une plage et les copie dans une autre plage.

```js
Excel.run(function (ctx) { 
    var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");
    ctx.load(range, "values");
    return ctx.sync().then(function() {
        var myvalues=range.values;
        ctx.workbook.worksheets.getActiveWorksheet().getRange("B1:B2").values = myvalues;
        console.log(range.values);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
})
```
