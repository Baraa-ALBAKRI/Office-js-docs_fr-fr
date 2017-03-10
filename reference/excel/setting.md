# <a name="setting-object-javascript-api-for-excel"></a>Objet Setting (API JavaScript pour Excel)

Setting représente une paire clé-valeur d’un paramètre conservé dans le document.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|Key|chaîne|Renvoie la clé qui représente l’ID du paramètre. En lecture seule.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|value|object|Représente la valeur stockée pour ce paramètre.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Supprime le paramètre.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="delete"></a>delete()
Supprime le paramètre.

#### <a name="syntax"></a>Syntaxe
```js
settingObject.delete();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void
