# <a name="inkanalysisparagraphcollection-object-(javascript-api-for-onenote)"></a>Objet InkAnalysisParagraphCollection (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_  


Représente une collection d’objets InkAnalysisParagraph.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|count|int|Renvoie le nombre d’objets InkAnalysisParagraph dans la page. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-count)|
|items|[InkAnalysisParagraph[]](inkanalysisparagraph.md)|Collection d’objets inkAnalysisParagraph. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-items)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[InkAnalysisParagraph](inkanalysisparagraph.md)|Obtient un objet InkAnalysisParagraph en fonction de son ID ou de son index dans la collection. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkAnalysisParagraph](inkanalysisparagraph.md)|Obtient un objet InkAnalysisParagraph en fonction de sa position dans la collection.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-load)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getitem(index:-number-or-string)"></a>getItem(index: number or string)
Obtient un objet InkAnalysisParagraph en fonction de son ID ou de son index dans la collection. En lecture seule.

#### <a name="syntax"></a>Syntaxe
```js
inkAnalysisParagraphCollectionObject.getItem(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number or string|ID de l’objet InkAnalysisParagraph ou emplacement d’index de l’objet InkAnalysisParagraph dans la collection.|

#### <a name="returns"></a>Retourne
[InkAnalysisParagraph](inkanalysisparagraph.md)

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
Obtient un objet InkAnalysisParagraph en fonction de sa position dans la collection.

#### <a name="syntax"></a>Syntaxe
```js
inkAnalysisParagraphCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[InkAnalysisParagraph](inkanalysisparagraph.md)

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
