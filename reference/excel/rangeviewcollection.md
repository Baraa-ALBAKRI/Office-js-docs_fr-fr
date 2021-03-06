# <a name="rangeviewcollection-object-javascript-api-for-excel"></a>Objet RangeViewCollection (API JavaScript pour Excel)

Représente une collection d’objets RangeView.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|éléments|[RangeView[]](rangeview.md)|Collection d’objets rangeView. En lecture seule.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Obtient le nombre d’objets RangeView dans la collection.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[RangeView](rangeview.md)|Obtient une ligne d’affichage de plage par l’intermédiaire de son index. Avec index de base zéro.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getcount"></a>getCount()
Obtient le nombre d’objets RangeView dans la collection.

#### <a name="syntax"></a>Syntaxe
```js
rangeViewCollectionObject.getCount();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Renvoie
int

### <a name="getitematindex-number"></a>getItemAt(index: number)
Obtient une ligne d’affichage de plage par l’intermédiaire de son index. Avec index de base zéro.

#### <a name="syntax"></a>Syntaxe
```js
rangeViewCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|index|number|Index de la ligne visible.|

#### <a name="returns"></a>Retourne
[RangeView](rangeview.md)
