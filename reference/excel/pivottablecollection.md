# <a name="pivottablecollection-object-javascript-api-for-excel"></a>Objet PivotTableCollection (API JavaScript pour Excel)

Représente une collection de tous les tableaux croisés dynamiques du classeur ou de la feuille de travail.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|éléments|[PivotTable[]](pivottable.md)|Collection d’objets de tableau croisé dynamique. En lecture seule.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Obtient le nombre de tableaux croisés dynamiques de la collection.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[PivotTable](pivottable.md)|Extrait un tableau croisé dynamique par nom.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[PivotTable](pivottable.md)|Extrait un tableau croisé dynamique par nom. Si le tableau croisé dynamique n’existe pas, renvoie un objet null.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[refreshAll()](#refreshall)|void|Actualise tous les tableaux croisés dynamiques de la collection.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getcount"></a>getCount()
Obtient le nombre de tableaux croisés dynamiques de la collection.

#### <a name="syntax"></a>Syntaxe
```js
pivotTableCollectionObject.getCount();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Renvoie
int

### <a name="getitemname-string"></a>getItem(name: string)
Extrait un tableau croisé dynamique par nom.

#### <a name="syntax"></a>Syntaxe
```js
pivotTableCollectionObject.getItem(name);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|name|chaîne|Nom du tableau croisé dynamique à récupérer.|

#### <a name="returns"></a>Retourne
[PivotTable](pivottable.md)

### <a name="getitemornullobjectname-string"></a>getItemOrNullObject(name: string)
Extrait un tableau croisé dynamique par nom. Si le tableau croisé dynamique n’existe pas, renvoie un objet null.

#### <a name="syntax"></a>Syntaxe
```js
pivotTableCollectionObject.getItemOrNullObject(name);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|name|chaîne|Nom du tableau croisé dynamique à récupérer.|

#### <a name="returns"></a>Retourne
[PivotTable](pivottable.md)

### <a name="refreshall"></a>refreshAll()
Actualise tous les tableaux croisés dynamiques de la collection.

#### <a name="syntax"></a>Syntaxe
```js
pivotTableCollectionObject.refreshAll();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void
