
# <a name="labs.components.choicecomponentresult"></a>Labs.Components.ChoiceComponentResult

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Résultat de l’envoi d’un composant de choix.

```
class ChoiceComponentResult
```


## <a name="properties"></a>Propriétés


|Propriété|Description|
|:-----|:-----|
| `public var score: any`|Note associée à l’envoi.|
| `public var complete: boolean`|Indique si le résultat a mis fin à la tentative.  Indique **True** si le résultat a mis fin à la tentative.|

## <a name="methods"></a>Méthodes




### <a name="constructor"></a>constructeur

 `function constructor(score: any, complete: boolean)`

Crée une instance de la classe **ChoiceComponentResult**.

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _score_|Note du résultat.|
| _complete_|Indique si le résultat a mis fin à la tentative.|
