# <a name="chartpointformat-object-javascript-api-for-excel"></a>Objet ChartPointFormat (interface API JavaScript pour Excel)

Représente le formatage des points du graphique.

## <a name="properties"></a>Propriétés

Aucun

## <a name="relationships"></a>Relations
| Relation | Type   |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|remplissage|[ChartFill](chartfill.md)|Représente le format de remplissage d’un graphique, qui comprend les informations de mise en forme d’arrière-plan. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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
