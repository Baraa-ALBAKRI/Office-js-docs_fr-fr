# <a name="worksheetprotectionoptions-object-(javascript-api-for-excel)"></a>Objet WorksheetProtectionOptions (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Excel pour iOS, Office 2016_

Représente les options de protection d’une feuille de calcul.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|allowAutoFilter|bool|Représente l’option de protection de feuille de calcul qui autorise l’utilisation de la fonctionnalité Filtre automatique.|
|allowDeleteColumns|bool|Représente l’option de protection de feuille de calcul qui autorise la suppression des colonnes.|
|allowDeleteRows|bool|Représente l’option de protection de feuille de calcul qui autorise la suppression des lignes.|
|allowFormatCells|bool|Représente l’option de protection de feuille de calcul qui autorise la mise en forme des cellules.|
|allowFormatColumns|bool|Représente l’option de protection de feuille de calcul qui autorise la mise en forme des colonnes.|
|allowFormatRows|bool|Représente l’option de protection de feuille de calcul qui autorise la mise en forme des lignes.|
|allowInsertColumns|bool|Représente l’option de protection de feuille de calcul qui autorise l’insertion des colonnes.|
|allowInsertHyperlinks|bool|Représente l’option de protection de feuille de calcul qui autorise l’insertion des liens hypertexte.|
|allowInsertRows|bool|Représente l’option de protection de feuille de calcul qui autorise l’insertion des lignes.|
|allowPivotTables|bool|Représente l’option de protection de feuille de calcul qui autorise l’utilisation de la fonctionnalité Tableau croisé dynamique.|
|allowSort|bool|Représente l’option de protection de feuille de calcul qui autorise l’utilisation de la fonctionnalité Tri.|

_Voir des [exemples d’accès aux propriétés.](#examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

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
Cet exemple charge les options de protection de la feuille de calcul active.
```js
Excel.run(function (ctx) {
    var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
    worksheet.protection.load();            
    return ctx.sync()
        .then(function () {
            console.log("Active worksheet's protection options: " + worksheet.protection.options);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
