
# <a name="labs.components.activitycomponentinstance"></a>Labs.Components.ActivityComponentInstance

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Représente l’instance actuelle d’un composant d’activité.

```
class ActivityComponentInstance extends Labs.ComponentInstance<Components.ActivityComponentAttempt>
```


## <a name="properties"></a>Propriétés


|**Nom**|**Description**|
|:-----|:-----|
| `public var component: Components.IActivityComponentInstance`|Instance du composant d’activité [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md) sous-jacent représenté par cette classe.|

## <a name="methods"></a>Méthodes




### <a name="constructor"></a>constructeur

 `function constructor(component: Components.IActivityComponentInstance)`

Crée une instance de la classe [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md).

 **Paramètres**


|**Nom**|**Description**|
|:-----|:-----|
| _component_|Instance de composant **IActivityComponentInstance** permettant de créer cette classe à partir de cette classe.|

### <a name="buildattempt"></a>buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ActivityComponentAttempt`

Génère une instance **ActivityComponentAttempt** et implémente la méthode abstraite définie dans la classe de base.

 **Paramètres**


|**Nom**|**Description**|
|:-----|:-----|
| _createAttemptResult_|Résultat d’une tentative de création.|
