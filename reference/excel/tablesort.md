# <a name="tablesort-object-javascript-api-for-excel"></a>Objet TableSort (API JavaScript pour Excel)

Gère les opérations de tri des objets Table.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|matchCase|bool|Indique si la casse a influé sur le dernier tri du tableau. En lecture seule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|méthode|string|Dernière méthode de classement des caractères chinois utilisée pour trier le tableau. En lecture seule. Les valeurs possibles sont les suivantes : PinYin, StrokeCount|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="relationships"></a>Relations
| Relation | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|champs|[SortField](sortfield.md)|Dernières conditions utilisées pour trier le tableau. En lecture seule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[apply(fields: SortField[], matchCase: bool, method: string)](#applyfields-sortfield-matchcase-bool-method-string)|void|Effectue une opération de tri.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[clear()](#clear)|void|Efface le tri actuellement appliqué au tableau. Même si le classement du tableau n’est pas modifié, l’état des boutons d’en-tête est rétabli.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[reapply()](#reapply)|void|Applique à nouveau les paramètres actuels de tri au tableau.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="applyfields-sortfield-matchcase-bool-method-string"></a>apply(fields: SortField[], matchCase: bool, method: string)
Effectue une opération de tri.

#### <a name="syntax"></a>Syntaxe
```js
tableSortObject.apply(fields, matchCase, method);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|champs|SortField[]|Liste des conditions de tri.|
|matchCase|bool|Facultatif. Indique si la casse influe sur le classement des chaînes.|
|méthode|string|Facultatif. Méthode de classement utilisée pour les caractères chinois.  Les valeurs possibles sont les suivantes : PinYin, StrokeCount|

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.apply([ 
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

### <a name="clear"></a>clear()
Efface le tri actuellement appliqué au tableau. Même si le classement du tableau n’est pas modifié, l’état des boutons d’en-tête est rétabli.

#### <a name="syntax"></a>Syntaxe
```js
tableSortObject.clear();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

### <a name="reapply"></a>reapply()
Applique à nouveau les paramètres actuels de tri au tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableSortObject.reapply();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void
