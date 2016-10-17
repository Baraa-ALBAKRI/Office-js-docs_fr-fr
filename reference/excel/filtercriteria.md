# <a name="filtercriteria-object-(javascript-api-for-excel)"></a>Objet FilterCriteria (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Excel pour iOS, Office 2016_

Représente les critères de filtrage appliqués à une colonne.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|color|string|Chaîne de couleur HTML utilisée pour filtrer des cellules. Utilisée avec le filtrage « cellColor » et « fontColor ».|
|criterion1|string|Premier critère utilisé pour filtrer des données. Utilisé comme opérateur dans le cas d’un filtrage « Custom ».|
|criterion2|string|Second critère utilisé pour filtrer des données. Utilisé uniquement comme opérateur dans le cas d’un filtrage « Custom ».|
|dynamicCriteria|string|Critères dynamiques de l’ensemble Excel.DynamicFilterCriteria à appliquer à cette colonne. Utilisé avec un filtrage « Dynamic ». Les valeurs possibles sont les suivantes : Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|
|filterOn|chaîne|Propriété utilisée par le filtre pour déterminer si les valeurs doivent rester visibles. Les valeurs possibles sont les suivantes :    BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom |
|values|object[]|Valeurs à utiliser pour le filtrage « Values ».|

## <a name="relationships"></a>Relations
| Relation | Type   |Description|
|:---------------|:--------|:----------|
|icône|[Icon](icon.md)|Icône utilisée pour filtrer des cellules. Utilisé avec le filtrage « Icon ».|
|opérateur|[FilterOperator](filteroperator.md)|Opérateur utilisé pour combiner les critères 1 et 2 lorsque vous utilisez le filtrage « Custom ».|

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
