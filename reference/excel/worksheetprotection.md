# <a name="worksheetprotection-object-javascript-api-for-excel"></a>Objet WorksheetProtection (interface API JavaScript pour Excel)

Cet objet représente la protection d’un objet de la feuille.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|protégé|bool|Indique si la feuille de calcul est protégée. En lecture seule. En lecture seule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="relationships"></a>Relations
| Relation | Type   |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|options|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|Options de protection de feuille. En lecture seule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[protect(options: WorksheetProtectionOptions)](#protectoptions-worksheetprotectionoptions)|void|Protège une feuille de calcul. Échoue si la feuille de calcul est protégée.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[unprotect()](#unprotect)|void|Annule la protection d’une feuille de calcul.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


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

### <a name="protectoptions-worksheetprotectionoptions"></a>protect(options: WorksheetProtectionOptions)
Protège une feuille de calcul. Échoue si la feuille de calcul est protégée.

#### <a name="syntax"></a>Syntaxe
```js
worksheetProtectionObject.protect(options);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|options|WorksheetProtectionOptions|Facultatif. Options de protection de feuille.|

#### <a name="returns"></a>Renvoie
void

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    var range = sheet.getRange("A1:B3").format.protection.locked = false;
    sheet.protection.protect({allowInsertRows:true});
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});

```
### <a name="unprotect"></a>unprotect()
Annule la protection d’une feuille de calcul.

#### <a name="syntax"></a>Syntaxe
```js
worksheetProtectionObject.unprotect();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void
