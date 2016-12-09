# <a name="settingcollection-object-javascript-api-for-excel"></a>Objet SettingCollection (API JavaScript pour Excel)

Représente une collection d’objets de feuille de calcul qui font partie du classeur.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|éléments|[Setting[]](setting.md)|Collection d’objets setting. En lecture seule.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[getItem(key: string)](#getitemkey-string)|[Setting](setting.md)|Obtient une entrée Setting via la clé.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(key: chaîne)](#getitemornullkey-string)|[Setting](setting.md)|Obtient une entrée Setting via la clé. Si l’objet Setting n’existe pas, la propriété isNull de l’objet renvoyé aura la valeur true.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[set(key: chaîne, value: chaîne)](#setkey-string-value-string)|[Setting](setting.md)|Définit ou ajoute le paramètre spécifié dans le classeur.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getitemkey-string"></a>getItem(key: string)
Obtient une entrée Setting via la clé.

#### <a name="syntax"></a>Syntaxe
```js
settingCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|Key|chaîne|Clé du paramètre.|

#### <a name="returns"></a>Retourne
[Setting](setting.md)

### <a name="getitemornullkey-string"></a>getItemOrNull(key: chaîne)
Obtient une entrée Setting via la clé. Si l’objet Setting n’existe pas, la propriété isNull de l’objet renvoyé aura la valeur true.

#### <a name="syntax"></a>Syntaxe
```js
settingCollectionObject.getItemOrNull(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|Key|chaîne|Clé du paramètre.|

#### <a name="returns"></a>Retourne
[Setting](setting.md)

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

### <a name="setkey-string-value-string"></a>set(key: chaîne, value: chaîne)
Définit ou ajoute le paramètre spécifié dans le classeur.

#### <a name="syntax"></a>Syntaxe
```js
settingCollectionObject.set(key, value);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|Key|chaîne|Clé du nouveau paramètre.|
|value|chaîne|Valeur du nouveau paramètre.|

#### <a name="returns"></a>Retourne
[Setting](setting.md)
