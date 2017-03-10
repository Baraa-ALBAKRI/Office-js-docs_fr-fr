# <a name="pageloadcompleteeventargs-object-javascript-api-for-visio"></a>Objet PageLoadCompleteEventArgs (API JavaScript pour Visio)

S’applique à : _Visio Online_

Fournit des informations sur la page qui a déclenché l’événement PageLoadComplete.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description
|:---------------|:--------|:----------|
|pageName|string|Obtient le nom de la page qui a déclenché l’événement PageLoad.|
|success|bool|Obtient le succès ou l’échec de l’événement PageLoadComplete.|

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
             eventResult1 = document1.onPageLoadComplete.add(
            function (args){
                   console.log("Page name: "+args.pageName);
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
