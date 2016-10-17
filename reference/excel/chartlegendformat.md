# <a name="chartlegendformat-object-(javascript-api-for-excel)"></a>Objet ChartLegendFormat (interface API JavaScript pour Excel)

Regroupe les propriétés de format de la légende d’un graphique.

## <a name="properties"></a>Propriétés

Aucun

## <a name="relationships"></a>Relations
| Relation | Type   |Description|
|:---------------|:--------|:----------|
|remplissage|[ChartFill](chartfill.md)|Représente le format de remplissage d’un objet, qui comprend des informations de mise en forme d’arrière-plan. En lecture seule.|
|police|[ChartFont](chartfont.md)|Représente les attributs de police, comme le nom de la police, la taille de police, la couleur, etc. pour la légende d’un graphique. En lecture seule.|

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
