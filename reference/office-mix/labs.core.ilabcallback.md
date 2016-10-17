
# <a name="labs.core.ilabcallback"></a>Labs.Core.ILabCallback

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Interface de gestion des méthodes de rappel Labs.js.

```
interface ILabCallback<T>
```


## <a name="callback-signature"></a>Signature de rappel

 `(err: any, data: T): void`

 **Paramètres de rappel**


|||
|:-----|:-----|
| _err_|**Null** si aucune erreur ne se produit. Autre réponse que **null** si une erreur s’est produite.|
| _data_|Données renvoyées avec le rappel.|
