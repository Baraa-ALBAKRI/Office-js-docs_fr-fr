# <a name="application-object-javascript-api-for-excel"></a>Objet Application (interface API JavaScript pour Excel)

Représente l’application Excel qui gère le classeur.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description|Dem. Set|
|:---------------|:--------|:----------|:----------|
|calculationMode|string|Renvoie le mode de calcul du classeur. En lecture seule. Les valeurs possibles sont les suivantes : `Automatic` Excel contrôle le recalcul, `AutomaticExceptTables` Excel contrôle le recalcul, mais ignore les modifications apportées aux tables, `Manual` le calcul est effectué lorsque l’utilisateur le demande.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|Dem. Set|
|:---------------|:--------|:----------|:----------|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|Recalcule tous les classeurs actuellement ouverts dans Excel.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="calculatecalculationtype-string"></a>calculate(calculationType: string)
Recalcule tous les classeurs actuellement ouverts dans Excel.

#### <a name="syntax"></a>Syntaxe
```js
applicationObject.calculate(calculationType);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|calculationType|string|Spécifie le type de calcul à utiliser. Les valeurs possibles sont les suivantes : `Recalculate` Ceci est un recalcul arrondi, principalement effectué pour des raisons de compatibilité descendante. `Full` Recalcule toutes les cellules qu’Excel a marquées comme compromises, c'est-à-dire celles dépendant de données volatiles ou modifiées, et les cellules marquées comme compromises par programmation. `FullRebuild` Recalcule toutes les cellules de tous les classeurs ouverts.|

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    ctx.workbook.application.calculate('Full');
    return ctx.sync(); 
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
|:---------------|:--------|:----------|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, accepte un objet [loadOption](loadoption.md).|

#### <a name="returns"></a>Renvoie
void
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés
```js
Excel.run(function (ctx) { 
    var application = ctx.workbook.application;
    application.load('calculationMode');
    return ctx.sync().then(function() {
        console.log(application.calculationMode);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
