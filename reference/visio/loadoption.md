# <a name="object-load-options-javascript-api-for-visio"></a>Objet Load Options (interface API JavaScript pour Visio)

Représente un objet qui peut être transmis à la méthode de chargement pour spécifier l’ensemble des propriétés et des relations à charger lors de l’exécution de la méthode **sync()** qui synchronise les états entre les objets Visio et les objets de proxy JavaScript correspondants. Cet objet utilise des options telles que les paramètres des propriétés select et expand pour spécifier l’ensemble des propriétés à charger sur l’objet et autorise également la pagination sur la collection.

Vous pouvez également fournir une chaîne ou un tableau qui contient les propriétés et les relations à charger, tel qu’illustré dans l’exemple suivant.

```js
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

## <a name="properties"></a>Propriétés

| Propriété | Type  | Description |
|:---------|:------|:------------|
|select    |object |Fournissez un tableau ou une liste de noms de paramètres/relations (en les séparant par des virgules) à charger lors de l’appel de la méthode executeAsync. Par exemple, "propriété1, relation1", ["propriété1", "relation1"]. Facultatif.|
|expand    |object |Fournissez un tableau ou une liste de noms de relations (en les séparant par des virgules) à charger lors de l’appel de la méthode executeAsync. Par exemple, "relation1, relation2", [ "relation1", "relation2"]. Facultatif.|
|top       |int    |Indiquez le nombre d’éléments de la collection demandée à inclure dans le résultat. Facultatif.|
|skip      |int    |Indiquez le nombre d’éléments de la collection devant être ignorés et exclus du résultat. Si une valeur est définie pour **top**, la sélection du résultat démarre une fois que le nombre spécifié d’éléments a été ignoré. Facultatif.|

