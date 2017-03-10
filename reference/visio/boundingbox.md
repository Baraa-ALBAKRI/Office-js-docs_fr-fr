# <a name="boundingbox-object-javascript-api-for-visio"></a>Objet BoundingBox (API JavaScript pour Visio)

S’applique à : _Visio Online_

Représente le BoundingBox de la forme.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description|
|:---------------|:--------|:----------|
|height|int|Distance entre les bords supérieur et inférieur du cadre englobant de la forme, à l’exclusion des graphiques de données associées à la forme.|
|width|int|Distance entre les côtés gauche et droit du cadre englobant de la forme, à l’exclusion des graphiques de données associés à la forme.|
|x|int|Nombre entier indiquant l’axe des abscisses du cadre englobant.|
|y|int|Nombre entier indiquant l’axe des ordonnées du cadre englobant.|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## <a name="method-details"></a>Détails des méthodes


### <a name="loadparam-object"></a>load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### <a name="syntax"></a>Syntaxe
```js
object.load(param);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### <a name="returns"></a>Renvoie
void
