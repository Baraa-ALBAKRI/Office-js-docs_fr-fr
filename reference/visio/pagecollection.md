# <a name="pagecollection-object-javascript-api-for-visio"></a>Objet PageCollection (interface API JavaScript pour Visio)

S’applique à : _Visio Online_

Représente une collection d’objets Page faisant partie du document.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description|
|:---------------|:--------|:----------|
|éléments|[Page[]](page.md)|Collection d’objets de page. En lecture seule.|

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[getCount()](#getcount)|int|Obtient le nombre de pages de la collection.|
|[getItem(key: number ou string)](#getitemkey-number-or-string)|[Page](page.md)|Obtient une page à l’aide de sa clé (nom ou ID).|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## <a name="method-details"></a>Détails des méthodes


### <a name="getcount"></a>getCount()
Obtient le nombre de pages de la collection.

#### <a name="syntax"></a>Syntaxe
```js
pageCollectionObject.getCount();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Renvoie
int

### <a name="getitemkey-number-or-string"></a>getItem(key: valeur numérique ou chaîne)
Obtient une page à l’aide de sa clé (nom ou ID).

#### <a name="syntax"></a>Syntaxe
```js
pageCollectionObject.getItem(key);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|Key|valeur numérique ou chaîne|La clé est le nom ou l’ID de la page à récupérer.|

#### <a name="returns"></a>Renvoie
[Page](page.md)

#### <a name="examples"></a>Exemples
```js
Visio.run(function (ctx) { 
    var pageName = 'Page-1';
    var page = ctx.document.pages.getItem(pageName);
    page.activate();
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

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
