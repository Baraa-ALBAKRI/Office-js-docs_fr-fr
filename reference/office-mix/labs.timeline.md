
# <a name="labs.timeline"></a>Labs.Timeline

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Fournit un accès à la fonctionnalité de chronologie labs.js.

```
class Timeline
```


## <a name="methods"></a>Méthodes




### <a name="method"></a>méthode

 `function constructor(labsInternal: Labs.LabsInternal)`

Crée une instance de la classe **Timeline**.


### <a name="next"></a>suivant

 `public function next(completionStatus: Labs.Core.ICompletionStatus, callback: Labs.Core.ILabCallback<void>): void`

Indique que la chronologie doit passer à la diapositive suivante.

 **Paramètres**


|||
|:-----|:-----|
| _completionStatus_|Indique l’état actuel de l’atelier.|
| _callback_|Fonction de rappel qui se déclenche quand l’atelier passe à la diapositive suivante.|
