# <a name="workbook-object-(javascript-api-for-excel)"></a>Objet Workbook (interface API JavaScript pour Excel)

Le classeur est l’objet de niveau supérieur qui contient des objets connexes tels que des feuilles de calcul, des tableaux, des plages, etc.

## <a name="properties"></a>Propriétés

Aucun

## <a name="relationships"></a>Relations
| Relation | Type   |Description|
|:---------------|:--------|:----------|
|application|[Application](application.md)|Représente une instance de l’application Excel contenant ce classeur. En lecture seule.|
|liaisons|[BindingCollection](bindingcollection.md)|Représente une collection de liaisons appartenant au classeur. En lecture seule.|
|fonctions|[Functions](functions.md)|Représente l’instance de l’application Excel contenant ce classeur. En lecture seule.|
|noms|[NamedItemCollection](nameditemcollection.md)|Représente une collection d’éléments nommés portant sur le classeur (appelés plages et constantes). En lecture seule.|
|tableaux|[TableCollection](tablecollection.md)|Représente une collection de tableaux associés au classeur. En lecture seule.|
|feuilles de calcul|[WorksheetCollection](worksheetcollection.md)|Représente une collection de feuilles de calcul associées au classeur. En lecture seule.|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|Obtient la plage sélectionnée dans le classeur.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## <a name="method-details"></a>Détails des méthodes


### <a name="getselectedrange()"></a>getSelectedRange()
Obtient la plage sélectionnée dans le classeur.

#### <a name="syntax"></a>Syntaxe
```js
workbookObject.getSelectedRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples

```js
Excel.run(function (ctx) { 
    var selectedRange = ctx.workbook.getSelectedRange();
    selectedRange.load('address');
    return ctx.sync().then(function() {
            console.log(selectedRange.address);
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
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### <a name="returns"></a>Renvoie
void
