# <a name="pivottablecollection-object-javascript-api-for-excel"></a>Objet PivotTableCollection (API JavaScript pour Excel)

Représente une collection de tous les tableaux croisés dynamiques du classeur ou de la feuille de travail.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|éléments|[PivotTable[]](pivottable.md)|Collection d’objets de tableau croisé dynamique. En lecture seule.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[getItem(name: string)](#getitemname-string)|[PivotTable](pivottable.md)|Extrait un tableau croisé dynamique par nom.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(name: chaîne)](#getitemornullname-string)|[PivotTable](pivottable.md)|Extrait un tableau croisé dynamique par nom. Si le tableau croisé dynamique n’existe pas, la propriété isNull de l’objet renvoyé aura la valeur true.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[refreshAll()](#refreshall)|void|Actualise tous les tableaux croisés dynamiques de la collection.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getitemname-string"></a>getItem(name: string)
Extrait un tableau croisé dynamique par nom.

#### <a name="syntax"></a>Syntaxe
```js
pivotTableCollectionObject.getItem(name);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|name|chaîne|Nom du tableau croisé dynamique à récupérer.|

#### <a name="returns"></a>Retourne
[PivotTable](pivottable.md)

### <a name="getitemornullname-string"></a>getItemOrNull(name: chaîne)
Extrait un tableau croisé dynamique par nom. Si le tableau croisé dynamique n’existe pas, la propriété isNull de l’objet renvoyé aura la valeur true.

#### <a name="syntax"></a>Syntaxe
```js
pivotTableCollectionObject.getItemOrNull(name);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|name|chaîne|Nom du tableau croisé dynamique à récupérer.|

#### <a name="returns"></a>Retourne
[PivotTable](pivottable.md)

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
