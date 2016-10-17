# <a name="worksheetprotection-object-(javascript-api-for-excel)"></a>Objet WorksheetProtection (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Excel pour iOS, Office 2016_

Représente la protection d’un objet de la feuille.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|protégé|bool|Indique si la feuille de calcul est protégée. En lecture seule.|

## <a name="relationships"></a>Relations
| Relation | Type   |Description|
|:---------------|:--------|:----------|
|options|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|Options de protection de feuille. En lecture seule.|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Insère les détails de protection de la feuille dans l'objet proxy.|
|[protect(options: WorksheetProtectionOptions)](#protectoptions-worksheetprotectionoption)|void|Protège une feuille de calcul. Générée si la feuille de calcul est protégée.|
|[unprotect()](#unprotect)|void|Ôte la protection d'une feuille de calcul|

## <a name="method-details"></a>Détails des méthodes


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

#### <a name="examples"></a>Exemples
Cet exemple charge les informations de protection de la feuille de calcul active.
```js
Excel.run(function (ctx) {
    var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
    worksheet.protection.load();            
    return ctx.sync()
        .then(function () {
            console.log("Active worksheet's protection status: " + worksheet.protection.protected);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="protect(options:-worksheetprotectionoptions)"></a>protect(options: WorksheetProtectionOptions)
Protège une feuille de calcul avec des stratégies de protection facultatives. Une exception est générée si la feuille de calcul est protégée. 

Lorsque des options sont spécifiées, des stratégies individuelles peuvent être activées ou désactivées. Si vous ne spécifiez aucune stratégie, une stratégie par défaut est activée. 

#### <a name="syntax"></a>Syntaxe
```js
worksheetProtectionObject.protect(options);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
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
### <a name="unprotect()"></a>unprotect()
Ôte la protection d'une feuille de calcul. 

#### <a name="syntax"></a>Syntaxe
```js
worksheetProtectionObject.unprotect();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");  
    sheet.protection.unprotect();
    return ctx.sync(); 
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```