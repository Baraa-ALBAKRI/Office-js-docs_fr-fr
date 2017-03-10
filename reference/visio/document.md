# <a name="document-object-javascript-api-for-visio"></a>Objet Document (interface API JavaScript pour Visio)

S’applique à : _Visio Online_

Représente la classe Document.

## <a name="properties"></a>Propriétés

Aucun

## <a name="relationships"></a>Relations
| Relation | Type    |Description|
|:---------------|:--------|:----------|
|application|[Application](application.md)|Représente une instance de l’application Visio contenant ce document. En lecture seule.|
|pages|[PageCollection](pagecollection.md)|Représente une collection de pages associées au document. En lecture seule.|
|view|[DocumentView](documentview.md)|Renvoie l’objet DocumentView. En lecture seule.|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[getActivePage()](#getactivepage)|[Page](page.md)|Renvoie la page active du document.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|
|[setActivePage(PageName: string)](#setactivepagepagename-string)|void|Configure la page active du document.|
|[startDataRefresh()](#startdatarefresh)|void|Déclenche l’actualisation des données dans le diagramme pour toutes les pages.|

## <a name="method-details"></a>Détails des méthodes


### <a name="getactivepage"></a>getActivePage()
Renvoie la page active du document.

#### <a name="syntax"></a>Syntaxe
```js
documentObject.getActivePage();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Page](page.md)

#### <a name="examples"></a>Exemples
```js
Visio.run(function (ctx) { 
    var document = ctx.document;
    var activePage = document.getActivePage();
    activePage.load();
    return ctx.sync().then(function () {
    console.log("pageName: " +activePage.name);
      });   
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

### <a name="setactivepagepagename-string"></a>setActivePage(PageName: chaîne)
Configure la page active du document.

#### <a name="syntax"></a>Syntaxe
```js
documentObject.setActivePage(PageName);
```

#### <a name="parameters"></a>Paramètres
| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|:---|
|PageName|chaîne|Nom de la page|

#### <a name="returns"></a>Renvoie
void

#### <a name="examples"></a>Exemples
```js
Visio.run(function (ctx) { 
    var document = ctx.document;
    var pageName = "Page-1";
    document.setActivePage(pageName);
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="startdatarefresh"></a>startDataRefresh()
Déclenche l’actualisation des données dans le diagramme pour toutes les pages.

#### <a name="syntax"></a>Syntaxe
```js
documentObject.startDataRefresh();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
Visio.run(function (ctx) { 
    var document = ctx.document;
    document.startDataRefresh();
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```### Property access examples
```js
Visio.run(function (ctx) { 
    var pages = ctx.document.pages;
    var pageCount = pages.getCount();
    return ctx.sync().then(function () {
        console.log("Pages Count: " +pageCount.value);
        });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>Exemples d’accès aux propriétés
```js
Visio.run(function (ctx) { 
    var documentView = ctx.document.view;
    documentView.disableHyperlinks();
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>Exemples d’accès aux propriétés
```js
Visio.run(function (ctx) { 
    var application = ctx.document.application;
    application.showToolbars = false;
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

