# <a name="datarefreshcompleteeventargs-object-javascript-api-for-visio"></a>Objet DataRefreshCompleteEventArgs (API JavaScript pour Visio)

S’applique à : _Visio Online_

Fournit des informations sur le document qui a déclenché l’événement DataRefreshComplete.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description
|:---------------|:--------|:----------|
|success|bool|Obtient le résultat successfailure de l’événement DataRefreshComplete.|
|document|[Document](document.md)|Obtient l’objet de document qui a déclenché l’événement DataRefreshComplete.|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

## <a name="relationships"></a>Relations
Aucun

## <a name="methods"></a>Méthodes
Aucun

### <a name="property-access-examples"></a>Exemples d’accès aux propriétés
```js
Visio.run(function (ctx) { 
  var document1= ctx.document;
               var page = document1.getActivePage();
         eventResult1 = document1.onDataRefreshComplete.add(
    function (args){
           console.log("Data Refresh Result: "+args.success);
    });

    return ctx.sync().then(function () {
       console.log("Success");
    });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
    console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
