# <a name="filtercriteria-object-javascript-api-for-excel"></a>Objet FilterCriteria (API JavaScript pour Excel)

Représente les critères de filtrage appliqués à une colonne.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|color|string|Chaîne de couleur HTML utilisée pour filtrer des cellules. Utilisée avec le filtrage « cellColor » et « fontColor ».|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|criterion1|string|Premier critère utilisé pour filtrer des données. Utilisé comme opérateur dans le cas d’un filtrage « Custom ».|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|criterion2|string|Second critère utilisé pour filtrer des données. Utilisé uniquement comme opérateur dans le cas d’un filtrage « Custom ».|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|dynamicCriteria|string|Critères dynamiques de l’ensemble Excel.DynamicFilterCriteria à appliquer à cette colonne. Utilisé avec un filtrage « Dynamic ». Les valeurs possibles sont les suivantes : Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|filterOn|chaîne|Propriété utilisée par le filtre pour déterminer si les valeurs doivent rester visibles. Les valeurs possibles sont les suivantes : BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|opérateur|chaîne|Opérateur utilisé pour combiner les critères 1 et 2 lorsque vous utilisez le filtrage « Custom ». Les valeurs possibles sont les suivantes : And, Or.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[]|Valeurs à utiliser pour le filtrage « Values ».|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
| Relation | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|icône|[Icon](icon.md)|Icône utilisée pour filtrer des cellules. Utilisé avec le filtrage « Icon ».|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes
Aucun

