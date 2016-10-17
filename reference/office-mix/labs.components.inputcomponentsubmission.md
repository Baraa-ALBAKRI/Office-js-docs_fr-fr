
# <a name="labs.components.inputcomponentsubmission"></a>Labs.Components.InputComponentSubmission

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Représente un envoi vers un composant de saisie.

```
class InputComponentSubmission
```


## <a name="properties"></a>Propriétés


|Propriété|Description|
|:-----|:-----|
| `public var answer: Components.InputComponentAnswer`|Réponse ([Labs.Components.InputComponentAnswer](../../reference/office-mix/labs.components.inputcomponentanswer.md)) associée à l’envoi.|
| `public var result: Components.InputComponentResult`|Résultat ([Labs.Components.InputComponentResult](../../reference/office-mix/labs.components.inputcomponentresult.md)) de l’envoi.|
| `public var time: number`|Heure à laquelle l’envoi a été reçu.|

## <a name="methods"></a>Méthodes




### <a name="constructor"></a>constructeur

 `function constructor(answer: Components.InputComponentAnswer, result: Components.InputComponentResult, time: number)`

Crée une instance de la classe **InputComponentSubmission**.

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _answer_|Réponse associée à l’envoi.|
| _result_|Résultat de l’envoi.|
| _time_|Heure à laquelle l’envoi a été reçu.|
