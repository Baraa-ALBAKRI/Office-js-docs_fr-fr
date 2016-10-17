# <a name="inkanalysiswordcollection-object-(javascript-api-for-onenote)"></a>Objet InkAnalysisWordCollection (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_  


Représente une collection d’objets InkAnalysisWord.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|count|int|Renvoie le nombre d’objets InkAnalysisWord dans la page. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-count)|
|items|[InkAnalysisWord[]](inkanalysisword.md)|Collection d’objets inkAnalysisWord. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-items)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[InkAnalysisWord](inkanalysisword.md)|Obtient un objet InkAnalysisWord en fonction de son ID ou de son index dans la collection. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkAnalysisWord](inkanalysisword.md)|Obtient un objet InkAnalysisWord sur sa position dans la collection de sites.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-load)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getitem(index:-number-or-string)"></a>getItem(index: number or string)
Obtient un objet InkAnalysisWord en fonction de son ID ou de son index dans la collection. En lecture seule.

#### <a name="syntax"></a>Syntaxe
```js
inkAnalysisWordCollectionObject.getItem(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number or string|ID de l’objet InkAnalysisWord ou emplacement d’index de l’objet InkAnalysisWord dans la collection.|

#### <a name="returns"></a>Retourne
[InkAnalysisWord](inkanalysisword.md)

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
Obtient un objet InkAnalysisWord sur sa position dans la collection de sites.

#### <a name="syntax"></a>Syntaxe
```js
inkAnalysisWordCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[InkAnalysisWord](inkanalysisword.md)

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

#### <a name="returns"></a>Retourne
void
