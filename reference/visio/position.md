# <a name="position-object-javascript-api-for-visio"></a>Objet Position (API JavaScript pour Visio)

S’applique à : _Visio Online_

Représente la position de l’objet dans l’affichage.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description|
|:---------------|:--------|:----------|
|x|int|Nombre entier spécifiant l’axe des abscisses de l’objet, qui est la valeur signée en pixels de la distance entre le milieu de la fenêtre d’affichage et le bord gauche de la page.|
|y|int|Nombre entier spécifiant l’axe des ordonnées de l’objet, qui est la valeur signée en pixels de la distance entre le milieu de la fenêtre d’affichage et le bord supérieur de la page.|

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
