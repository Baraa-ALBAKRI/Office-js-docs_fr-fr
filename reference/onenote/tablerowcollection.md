# <a name="tablerowcollection-object-(javascript-api-for-onenote)"></a>Objet TableRowCollection (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_  


Contient une collection d’objets TableRow.

## <a name="properties"></a>Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|count|int|Renvoie le nombre de lignes de tableau dans cette collection. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-count)|
|items|[TableRow[]](tablerow.md)|Collection d’objets tableRow. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-items)|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[TableRow](tablerow.md)|Obtient un objet de ligne de tableau en fonction de son ID ou de son index dans la collection. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[TableRow](tablerow.md)|Obtient une ligne de tableau au niveau de sa position dans la collection.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-load)|

## <a name="method-details"></a>Détails des méthodes


### <a name="getitem(index:-number-or-string)"></a>getItem(index: number or string)
Obtient un objet de ligne de tableau en fonction de son ID ou de son index dans la collection. En lecture seule.

#### <a name="syntax"></a>Syntaxe
```js
tableRowCollectionObject.getItem(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number or string|Nombre qui identifie l’emplacement associé à l’index d’un objet de ligne de tableau.|

#### <a name="returns"></a>Retourne
[TableRow](tablerow.md)

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
Obtient une ligne de tableau au niveau de sa position dans la collection.

#### <a name="syntax"></a>Syntaxe
```js
tableRowCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### <a name="returns"></a>Retourne
[TableRow](tablerow.md)

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
