# <a name="rangesort-object-(javascript-api-for-excel)"></a>Objet RangeSort (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Excel pour iOS, Office 2016_

Gère les opérations de tri des objets Range.

## <a name="properties"></a>Propriétés

Aucun

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)](#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|void|Effectue une opération de tri.|

## <a name="method-details"></a>Détails des méthodes


### <a name="apply(fields:-sortfield[],-matchcase:-bool,-hasheaders:-bool,-orientation:-string,-method:-string)"></a>apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)
Effectue une opération de tri.

#### <a name="syntax"></a>Syntaxe
```js
rangeSortObject.apply(fields, matchCase, hasHeaders, orientation, method);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|champs|SortField[]|Liste des conditions de tri.|
|matchCase|bool|Facultatif. Indique si la casse influe sur le classement des chaînes.|
|hasHeaders|bool|Facultatif. Indique si la plage comporte un en-tête.|
|orientation|string|Facultatif. Indique si l’opération trie les lignes ou les colonnes.  Les valeurs possibles sont les suivantes : Rows, Columns|
|méthode|string|Facultatif. Méthode de classement utilisée pour les caractères chinois.  Les valeurs possibles sont les suivantes : PinYin, StrokeCount|

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:G6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.sort.apply([ 
            {
                key: 2,
                ascending: true
            },
        ], true);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```