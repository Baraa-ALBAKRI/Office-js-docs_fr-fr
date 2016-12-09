# <a name="worksheetprotectionoptions-object-javascript-api-for-excel"></a>Objet WorksheetProtectionOptions (interface API JavaScript pour Excel)

Représente les options de protection d’une feuille de calcul.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|allowAutoFilter|bool|Représente l’option de protection de feuille de calcul qui autorise l’utilisation de la fonctionnalité Filtre automatique.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowDeleteColumns|bool|Représente l’option de protection de feuille de calcul qui autorise la suppression des colonnes.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowDeleteRows|bool|Représente l’option de protection de feuille de calcul qui autorise la suppression des lignes.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowFormatCells|bool|Représente l’option de protection de feuille de calcul qui autorise la mise en forme des cellules.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowFormatColumns|bool|Représente l’option de protection de feuille de calcul qui autorise la mise en forme des colonnes.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowFormatRows|bool|Représente l’option de protection de feuille de calcul qui autorise la mise en forme des lignes.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowInsertColumns|bool|Représente l’option de protection de feuille de calcul qui autorise l’insertion des colonnes.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowInsertHyperlinks|bool|Représente l’option de protection de feuille de calcul qui autorise l’insertion des liens hypertexte.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowInsertRows|bool|Représente l’option de protection de feuille de calcul qui autorise l’insertion des lignes.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowPivotTables|bool|Représente l’option de protection de feuille de calcul qui autorise l’utilisation de la fonctionnalité Tableau croisé dynamique.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowSort|bool|Représente l’option de protection de feuille de calcul qui autorise l’utilisation de la fonctionnalité Tri.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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
