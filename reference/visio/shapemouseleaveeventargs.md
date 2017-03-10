# <a name="shapemouseleaveeventargs-object-javascript-api-for-visio"></a>Objet ShapeMouseLeaveEventArgs (API JavaScript pour Visio)

S’applique à : _Visio Online_

Fournit des informations sur la forme qui a déclenché l’événement MouseLeave.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description
|:---------------|:--------|:----------|
|shapeName|string|Obtient le nom de l’objet de forme qui a déclenché l’événement MouseLeave.|
|pageName|string|Obtient le nom de la page comportant l’objet de forme qui a déclenché l’événement MouseLeave.|

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
    eventResult2 = document1.onMouseLeave.add(
                function (args){            
                         console.log(Date.now()+":OnMouseLeave Event"+JSON.stringify(args));
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