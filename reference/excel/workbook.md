# <a name="workbook-object-javascript-api-for-excel"></a>Objet Workbook (interface API JavaScript pour Excel)

Le classeur est l’objet de niveau supérieur qui contient des objets connexes tels que des feuilles de calcul, des tableaux, des plages, etc.

## <a name="properties"></a>Propriétés

Aucun

## <a name="relationships"></a>Relations
| Relation | Type   |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|application|[Application](application.md)|Représente l’instance de l’application Excel contenant ce classeur. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|bindings|[BindingCollection](bindingcollection.md)|Représente une collection de liaisons appartenant au classeur. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|fonctions|[Functions](functions.md)|Représente l’instance de l’application Excel contenant ce classeur. En lecture seule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|noms|[NamedItemCollection](nameditemcollection.md)|Représente une collection d’éléments nommés portant sur le classeur (appelés plages et constantes). En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|pivotTables|[PivotTableCollection](pivottablecollection.md)|Représente une collection de tableaux croisés dynamiques associés au classeur. En lecture seule.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|paramètres|[SettingCollection](settingcollection.md)|Représente une collection d’objets Settings associés au classeur. En lecture seule.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|tables|[TableCollection](tablecollection.md)|Représente une collection de tableaux associés au classeur. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|worksheets|[WorksheetCollection](worksheetcollection.md)|Représente une collection de feuilles de calcul associées au classeur. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|Obtient la plage sélectionnée dans le classeur.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getselectedrange"></a>getSelectedRange()
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
### <a name="loadparam-object"></a>load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### <a name="syntax"></a>Syntaxe
```js
object.load(param);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### <a name="returns"></a>Retourne
void
