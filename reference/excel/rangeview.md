# <a name="rangeview-object-javascript-api-for-excel"></a>Objet RangeView (API JavaScript pour Excel)

RangeView représente un ensemble de cellules visibles de la plage parent.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|cellAddresses|object[][]|Représente les adresses de cellule de la RangeView. En lecture seule.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|columnCount|int|Renvoie le nombre de colonnes visibles. En lecture seule.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulas|object[][]|Représente la formule dans le style de notation A1.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulasLocal|object[][]|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur. Par exemple, la formule « =SUM(A1, 1.5) » en anglais deviendrait « =SUMME(A1; 1,5) » en allemand.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulasR1C1|object[][]|Représente la formule dans le style de notation R1C1.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|index|int|Renvoie une valeur qui représente l’index de l’affichage de plage. En lecture seule.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|object[][]|Représente le code de format de nombre d’Excel pour une cellule donnée.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|int|Renvoie le nombre de lignes visibles. En lecture seule.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|text|object[][]|Valeurs de texte de la plage spécifiée. La valeur de texte ne dépend pas de la largeur de la cellule. Le remplacement par le signe # qui se produit dans l’interface utilisateur d’Excel n’a aucun effet sur la valeur de texte renvoyée par l’API. En lecture seule.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|valueTypes|string|Représente le type de données de chaque cellule. En lecture seule. Les valeurs possibles sont les suivantes : Unknown (inconnu), Empty (vide), String (chaîne), Integer (entier), Double (double), Boolean (valeur booléenne), Error (erreur).|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|Représente les valeurs brutes de l’affichage de plage spécifié. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Une cellule contenant une erreur renvoie la chaîne d’erreur.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type   |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|Objet Rows|[RangeViewCollection](rangeviewcollection.md)|Représente une collection d’affichages de plage associés à la plage. En lecture seule.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[getRange()](#getrange)|[Range](range.md)|Obtient la plage parent associée à l’affichage de plage actuel.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getrange"></a>getRange()
Obtient la plage parent associée à l’affichage de plage actuel.

#### <a name="syntax"></a>Syntaxe
```js
rangeViewObject.getRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

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
