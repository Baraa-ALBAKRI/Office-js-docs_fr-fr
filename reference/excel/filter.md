# <a name="filter-object-javascript-api-for-excel"></a>Objet Filter (interface API JavaScript pour Excel)

Gère le filtrage de la colonne d’un tableau.

## <a name="properties"></a>Propriétés

Aucun

## <a name="relationships"></a>Relations
| Relation | Type   |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|de réussite|[FilterCriteria](filtercriteria.md)|Le filtre actuellement appliqué à la colonne donnée. En lecture seule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[apply(criteria: FilterCriteria)](#applycriteria-filtercriteria)|void|Appliquer les critères de filtre donnés à la colonne indiquée.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyBottomItemsFilter(count: number)](#applybottomitemsfiltercount-number)|void|Appliquer un filtre « Élément inférieur » à la colonne pour le nombre d’éléments donné.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyBottomPercentFilter(percent: number)](#applybottompercentfilterpercent-number)|void|Appliquer un filtre « Pourcentage inférieur » à la colonne pour le pourcentage d’éléments donné.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyCellColorFilter(color: string)](#applycellcolorfiltercolor-string)|void|Appliquer un filtre « Couleur de cellule » à la colonne pour la couleur donnée.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyCustomFilter(criteria1: chaîne, criteria2: chaîne, oper: chaîne)](#applycustomfiltercriteria1-string-criteria2-string-oper-string)|void|Appliquer un filtre « Icône » à la colonne pour les chaînes de critères données.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyDynamicFilter(criteria: string)](#applydynamicfiltercriteria-string)|void|Appliquer un filtre « Dynamique » à la colonne.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyFontColorFilter(color: string)](#applyfontcolorfiltercolor-string)|void|Appliquer un filtre « Couleur de police » à la colonne pour la couleur donnée.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyIconFilter(icon: Icon)](#applyiconfiltericon-icon)|void|Appliquer un filtre « Icône » à la colonne pour l’icône donnée.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyTopItemsFilter(count: number)](#applytopitemsfiltercount-number)|void|Appliquer un filtre « Élément supérieur » à la colonne pour le nombre d’éléments donné.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyTopPercentFilter(percent: number)](#applytoppercentfilterpercent-number)|void|Appliquer un filtre « Pourcentage supérieur » à la colonne pour le pourcentage d’éléments donné.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyValuesFilter(values: ()[])](#applyvaluesfiltervalues-)|void|Appliquer un filtre « Valeurs » à la colonne pour les valeurs données.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[clear()](#clear)|void|Effacer le filtre sur la colonne donnée.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="applycriteria-filtercriteria"></a>apply(criteria: FilterCriteria)
Appliquer les critères de filtre donnés à la colonne indiquée.

#### <a name="syntax"></a>Syntaxe
```js
filterObject.apply(criteria);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|de réussite|FilterCriteria|Critères à appliquer.|

#### <a name="returns"></a>Renvoie
void

### <a name="applybottomitemsfiltercount-number"></a>applyBottomItemsFilter(count: number)
Appliquer un filtre « Élément inférieur » à la colonne pour le nombre d’éléments donné.

#### <a name="syntax"></a>Syntaxe
```js
filterObject.applyBottomItemsFilter(count);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|count|number|Nombre d’éléments à partir du bas à afficher.|

#### <a name="returns"></a>Renvoie
void

### <a name="applybottompercentfilterpercent-number"></a>applyBottomPercentFilter(percent: number)
Appliquer un filtre « Pourcentage inférieur » à la colonne pour le pourcentage d’éléments donné.

#### <a name="syntax"></a>Syntaxe
```js
filterObject.applyBottomPercentFilter(percent);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|pourcentage|number|Pourcentage d’éléments à partir du bas à afficher.|

#### <a name="returns"></a>Renvoie
void

### <a name="applycellcolorfiltercolor-string"></a>applyCellColorFilter(color: string)
Appliquer un filtre « Couleur de cellule » à la colonne pour la couleur donnée.

#### <a name="syntax"></a>Syntaxe
```js
filterObject.applyCellColorFilter(color);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|color|string|Couleur d’arrière-plan des cellules à afficher.|

#### <a name="returns"></a>Renvoie
void

### <a name="applycustomfiltercriteria1-string-criteria2-string-oper-string"></a>applyCustomFilter(criteria1: chaîne, criteria2: chaîne, oper: chaîne)
Appliquer un filtre « Icône » à la colonne pour les chaînes de critères données.

#### <a name="syntax"></a>Syntaxe
```js
filterObject.applyCustomFilter(criteria1, criteria2, oper);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|criteria1|string|Première chaîne de critères.|
|criteria2|string|Facultatif. Deuxième chaîne de critères.|
|oper|chaîne|Facultatif. Opérateur qui décrit comment les deux critères sont joints.  Les valeurs possibles sont les suivantes : And, Or|

#### <a name="returns"></a>Retourne
void

### <a name="applydynamicfiltercriteria-string"></a>applyDynamicFilter(criteria: string)
Appliquer un filtre « Dynamique » à la colonne.

#### <a name="syntax"></a>Syntaxe
```js
filterObject.applyDynamicFilter(criteria);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|de réussite|string|Critères dynamiques à appliquer.  Les valeurs possibles sont les suivantes : Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday|

#### <a name="returns"></a>Retourne
void

### <a name="applyfontcolorfiltercolor-string"></a>applyFontColorFilter(color: string)
Appliquer un filtre « Couleur de police » à la colonne pour la couleur donnée.

#### <a name="syntax"></a>Syntaxe
```js
filterObject.applyFontColorFilter(color);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|color|string|Couleur de police des cellules à afficher.|

#### <a name="returns"></a>Renvoie
void

### <a name="applyiconfiltericon-icon"></a>applyIconFilter(icon: Icon)
Appliquer un filtre « Icône » à la colonne pour l’icône donnée.

#### <a name="syntax"></a>Syntaxe
```js
filterObject.applyIconFilter(icon);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|icône|Icône|Icônes des cellules à afficher.|

#### <a name="returns"></a>Renvoie
void

### <a name="applytopitemsfiltercount-number"></a>applyTopItemsFilter(count: number)
Appliquer un filtre « Élément supérieur » à la colonne pour le nombre d’éléments donné.

#### <a name="syntax"></a>Syntaxe
```js
filterObject.applyTopItemsFilter(count);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|count|number|Nombre d’éléments à partir du haut à afficher.|

#### <a name="returns"></a>Renvoie
void

### <a name="applytoppercentfilterpercent-number"></a>applyTopPercentFilter(percent: number)
Appliquer un filtre « Pourcentage supérieur » à la colonne pour le pourcentage d’éléments donné.

#### <a name="syntax"></a>Syntaxe
```js
filterObject.applyTopPercentFilter(percent);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|pourcentage|number|Pourcentage d’éléments à partir du haut à afficher.|

#### <a name="returns"></a>Renvoie
void

### <a name="applyvaluesfiltervalues-"></a>applyValuesFilter(values: ()[])
Appliquer un filtre « Valeurs » à la colonne pour les valeurs données.

#### <a name="syntax"></a>Syntaxe
```js
filterObject.applyValuesFilter(values);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|values|()[]|Liste des valeurs à afficher.|

#### <a name="returns"></a>Renvoie
void

### <a name="clear"></a>clear()
Effacer le filtre sur la colonne donnée.

#### <a name="syntax"></a>Syntaxe
```js
filterObject.clear();
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
