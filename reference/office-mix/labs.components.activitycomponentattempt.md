
# <a name="labs.components.activitycomponentattempt"></a>Labs.Components.ActivityComponentAttempt

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Représente une tentative pour terminer un composant d’activité.

```
class Permissions
```


## <a name="methods"></a>Méthodes




### <a name="constructor"></a>constructeur

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Crée une instance de la classe **ActivityComponentAttempt**.

 **Paramètres**


|**Nom**|**Description**|
|:-----|:-----|
| _labs_|Instances de l’atelier ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx)) associées au composant.|
| _componentId_|ID du composant associé à la tentative.|
| _attemptId_|ID de la tentative.|
| _values_|Valeurs éventuelles associées au composant.|

### <a name="complete"></a>Intégration

 `public function complete(callback: Labs.Core.ILabCallback<void>): void`

Indicateur signalant la fin de l’activité.

 **Paramètres**


|**Nom**|**Description**|
|:-----|:-----|
| _callback_|Fonction de rappel appelée une fois l’activité terminée.|

### <a name="processaction"></a>processAction

 `public function processAction(action: Labs.Core.IAction): void`

Fonction qui s’exécute sur les actions récupérées d’une tentative donnée et qui renseigne l’état de l’atelier.

 **Paramètres**


|**Nom**|**Description**|
|:-----|:-----|
| _action_|Instance de l’action ([Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md)).|
