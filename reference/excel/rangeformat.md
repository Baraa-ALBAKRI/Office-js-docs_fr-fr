# <a name="rangeformat-object-javascript-api-for-excel"></a>Objet RangeFormat (interface API JavaScript pour Excel)

Objet de format qui comprend les propriétés de police, de remplissage, de bordures, d’alignement, etc. de la plage.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|columnWidth|double|Obtient ou définit la largeur de toutes les colonnes de la plage. Si les largeurs de colonne ne sont pas uniformes, la valeur « null » est renvoyée.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|horizontalAlignment|string|Représente l’alignement horizontal de l’objet spécifié. Les valeurs possibles sont les suivantes : General (général), Left (gauche), Center (centré), Right (droit), Fill (remplir), Justify (justifié), CenterAcrossSelection (centré pour toute la sélection), Distributed (distribué).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowHeight|double|Obtient ou définit la hauteur de toutes les lignes de la plage. Si les hauteurs de lignes ne sont pas uniformes, la valeur « null » est renvoyée.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|verticalAlignment|string|Représente l’alignement vertical de l’objet spécifié. Les valeurs possibles sont les suivantes : Top (haut), Center (centré), Bottom (bas), Justify (justifié), Distributed (distribué).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|wrapText|bool|Indique si Excel renvoie le texte à la ligne dans l’objet. Une valeur null indique que les paramètres de retour à la ligne ne sont pas les mêmes sur l’ensemble de la plage.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type   |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|bordures|[RangeBorderCollection](rangebordercollection.md)|Collection d’objets de bordure qui s’appliquent à l’ensemble de la plage sélectionnée. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|remplissage|[RangeFill](rangefill.md)|Renvoie l’objet de remplissage défini sur la plage globale. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|police|[RangeFont](rangefont.md)|Renvoie l’objet de police défini sur la plage globale sélectionnée. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|protection|[FormatProtection](formatprotection.md)|Renvoie l’objet de protection du format pour une plage. En lecture seule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[autofitColumns()](#autofitcolumns)|void|Modifie la largeur des colonnes de la plage active pour obtenir le meilleur ajustement, en fonction des données présentes dans les colonnes.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[autofitRows()](#autofitrows)|void|Modifie la hauteur des lignes de la plage active pour obtenir le meilleur ajustement, en fonction des données présentes dans les colonnes.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="autofitcolumns"></a>autofitColumns()
Modifie la largeur des colonnes de la plage active pour obtenir le meilleur ajustement, en fonction des données présentes dans les colonnes.

#### <a name="syntax"></a>Syntaxe
```js
rangeFormatObject.autofitColumns();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

### <a name="autofitrows"></a>autofitRows()
Modifie la hauteur des lignes de la plage active pour obtenir le meilleur ajustement, en fonction des données présentes dans les colonnes.

#### <a name="syntax"></a>Syntaxe
```js
rangeFormatObject.autofitRows();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

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
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

L’exemple ci-dessous sélectionne toutes les propriétés de format de la plage. 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.load(["format/*", "format/fill", "format/borders", "format/font"]);
    return ctx.sync().then(function() {
        console.log(range.format.wrapText);
        console.log(range.format.fill.color);
        console.log(range.format.font.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

L’exemple ci-dessous définit le nom de la police et la couleur de remplissage, et place le texte. 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.wrapText = true;
    range.format.font.name = 'Times New Roman';
    range.format.fill.color = '0000FF';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

L’exemple suivant ajoute des bordures de grille autour de la plage.

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.borders.getItem('InsideHorizontal').style = 'Continuous';
    range.format.borders.getItem('InsideVertical').style = 'Continuous';
    range.format.borders.getItem('EdgeBottom').style = 'Continuous';
    range.format.borders.getItem('EdgeLeft').style = 'Continuous';
    range.format.borders.getItem('EdgeRight').style = 'Continuous';
    range.format.borders.getItem('EdgeTop').style = 'Continuous';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```