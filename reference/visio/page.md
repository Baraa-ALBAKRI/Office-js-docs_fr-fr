# <a name="page-object-javascript-api-for-visio"></a>Objet Page (interface API JavaScript pour Visio)

S’applique à : _Visio Online_

Représente la classe Page.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description|
|:---------------|:--------|:----------|
|height|int|Renvoie la hauteur de la page. En lecture seule.|
|Index|int|Index de l’objet Page. En lecture seule.|
|isBackground|bool|Indique s’il s’agit d’une page d’arrière-plan ou non. En lecture seule.|
|name|chaîne|Nom de la page. En lecture seule.|
|width|int|Renvoie la largeur de la page. En lecture seule.|

## <a name="relationships"></a>Relations
| Relation | Type    |Description|
|:---------------|:--------|:----------|
|comments|[CommentCollection](commentcollection.md)|Renvoie la collection de commentaires. En lecture seule.|
|shapes|[ShapeCollection](shapecollection.md)|Représente les formes de l’objet Page. En lecture seule.|
|view|[PageView](pageview.md)|Renvoie l’affichage de la page. En lecture seule.|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[activate()](#activate)|void|Définit la page comme la page active du document.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## <a name="method-details"></a>Détails des méthodes


### <a name="activate"></a>activate()
Définit la page comme la page active du document.

#### <a name="syntax"></a>Syntaxe
```js
pageObject.activate();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

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
