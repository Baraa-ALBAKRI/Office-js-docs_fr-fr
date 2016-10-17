# <a name="inkanalysislinecollection-object-(javascript-api-for-onenote)"></a>Objet InkAnalysisLineCollection (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_  


Représente une collection d’objets InkAnalysisLine.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|count|int|Renvoie le nombre d’objets InkAnalysisLine dans la page. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-count)|
|items|[InkAnalysisLine[]](inkanalysisline.md)|Collection d’objets inkAnalysisLine. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-items)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[InkAnalysisLine](inkanalysisline.md)|Obtient un objet InkAnalysisLine en fonction de son ID ou de son index dans la collection. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkAnalysisLine](inkanalysisline.md)|Obtient un objet InkAnalysisLine sur sa position dans la collection.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-load)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getitem(index:-number-or-string)"></a>getItem(index: number or string)
Obtient un objet InkAnalysisLine en fonction de son ID ou de son index dans la collection. En lecture seule.

#### <a name="syntax"></a>Syntaxe
```js
inkAnalysisLineCollectionObject.getItem(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number or string|ID de l’objet InkAnalysisLine ou emplacement d’index de l’objet InkAnalysisLine dans la collection.|

#### <a name="returns"></a>Retourne
[InkAnalysisLine](inkanalysisline.md)

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
Obtient un objet InkAnalysisLine sur sa position dans la collection.

#### <a name="syntax"></a>Syntaxe
```js
inkAnalysisLineCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[InkAnalysisLine](inkanalysisline.md)

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
