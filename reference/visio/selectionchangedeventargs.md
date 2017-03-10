# <a name="selectionchangedeventargs-object-javascript-api-for-visio"></a>Objet SelectionChangedEventArgs (API JavaScript pour Visio)

S’applique à : _Visio Online_

Fournit des informations sur la collection de formes qui a déclenché l’événement SelectionChanged.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description
|:---------------|:--------|:----------|
|shapeNames|string[]|Obtient le tableau des noms de forme qui a déclenché l’événement SelectionChanged.|
|pageName|string|Obtient le nom de la page qui comporte l’objet ShapeCollection qui a déclenché l’événement SelectionChanged.|

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
             eventResult1 = document1.onSelectionChanged.add(
        function (args){
                   console.log("Selected Shape Name: "+args.shapeNames[0]);
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