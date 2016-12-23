# <a name="page-object-javascript-api-for-visio"></a>Objet Page (interface API JavaScript pour Visio)

S’applique à : _Visio Online_
>**Remarque :** Les interfaces API JavaScript pour Visio sont actuellement affichées dans l’aperçu et peuvent être modifiées. Elles ne sont actuellement pas prises en charge dans les environnements de production.

Représente la classe Page.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description| Commentaires|
|:---------------|:--------|:----------|:---|
|index|int|Index de l’objet Page. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-index)|
|isBackground|bool|Indique s’il s’agit d’une page d’arrière-plan ou non. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-isBackground)|
|name|chaîne|Nom de la page. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-name)|

## <a name="relationships"></a>Relations
| Relation | Type   |Description| Commentaires|
|:---------------|:--------|:----------|:---|
|shapes|[ShapeCollection](shapecollection.md)|Représente les formes de l’objet Page. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-shapes)|
|view|[PageView](pageview.md)|Renvoie l’affichage de la page. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-view)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:---|
|[activate()](#activate)|void|Définit la page comme la page active du document.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-activate)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-load)|

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
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|:---|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### <a name="returns"></a>Renvoie
void
