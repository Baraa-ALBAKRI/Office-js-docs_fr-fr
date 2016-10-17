
# <a name="labs.components.inputcomponentresult"></a>Labs.Components.InputComponentResult

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Résultat d’un envoi de composant de saisie.

```
class InputComponentResult
```


## <a name="properties"></a>Propriétés


|Propriété|Description|
|:-----|:-----|
| `public var score: any`|Note associée à l’envoi.|
| `public var complete: boolean`|Indique si le résultat envoyé a mis fin à la tentative.  Indique **True** si la tentative est terminée.|

## <a name="methods"></a>Méthodes




### <a name="constructor"></a>constructeur

 `function constructor(score: any, complete: boolean)`

Crée une instance de la classe **InputComponentResult**.

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _score_|Note associée au résultat.|
| _complete_|Indique l’expression booléenne **true** si le résultat a mis fin à la tentative.|
