# <a name="settingcollection-object-javascript-api-for-excel"></a>Objet SettingCollection (API JavaScript pour Excel)

Représente une collection d’objets de feuille de calcul qui font partie du classeur.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|éléments|[Setting[]](setting.md)|Collection d’objets setting. En lecture seule.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[add(key: string, value: (any)[])](#addkey-string-value-any)|[Setting](setting.md)|Définit ou ajoute le paramètre spécifié dans le classeur.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Obtient le nombre de paramètres dans la collection.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[Setting](setting.md)|Obtient une entrée de paramètre via la clé.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: string)](#getitemornullobjectkey-string)|[Setting](setting.md)|Obtient une entrée de paramètre via la clé. Si le paramètre n’existe pas, renvoie un objet null.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="addkey-string-value-any"></a>add(key: string, value: (any)[])
Définit ou ajoute le paramètre spécifié dans le classeur.

#### <a name="syntax"></a>Syntaxe
```js
settingCollectionObject.add(key, value);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|Key|chaîne|Clé du nouveau paramètre.|
|value|(any)[]|Valeur du nouveau paramètre.|

#### <a name="returns"></a>Retourne
[Setting](setting.md)

### <a name="getcount"></a>getCount()
Obtient le nombre de paramètres dans la collection.

#### <a name="syntax"></a>Syntaxe
```js
settingCollectionObject.getCount();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Renvoie
int

### <a name="getitemkey-string"></a>getItem(key: string)
Obtient une entrée Setting via la clé.

#### <a name="syntax"></a>Syntaxe
```js
settingCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|Key|chaîne|Clé du paramètre.|

#### <a name="returns"></a>Retourne
[Setting](setting.md)

### <a name="getitemornullobjectkey-string"></a>getItemOrNullObject(key: string)
Obtient une entrée de paramètre via la clé. Si le paramètre n’existe pas, renvoie un objet null.

#### <a name="syntax"></a>Syntaxe
```js
settingCollectionObject.getItemOrNullObject(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|Key|chaîne|Clé du paramètre.|

#### <a name="returns"></a>Retourne
[Setting](setting.md)
