
# <a name="labs.components.inputcomponentinstance"></a>Labs.Components.InputComponentInstance

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Représente une instance d’un composant de saisie.

```
class InputComponentInstance extends Labs.ComponentInstance<Components.InputComponentAttempt>
```


## <a name="properties"></a>Propriétés


|Propriété|Description|
|:-----|:-----|
| `public var component: Components.IInputComponentInstance`|Objet [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) sous-jacent représenté par cette classe.|

## <a name="methods"></a>Méthodes




### <a name="constructor"></a>constructeur

 `function constructor(component: Components.IInputComponentInstance)`

Crée une instance [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md).

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _component_|Objet [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) à utiliser pour créer cette classe.|

### <a name="buildattempt"></a>buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.InputComponentAttempt`

Crée une instance [Labs.Components.InputComponentAttempt](../../reference/office-mix/labs.components.inputcomponentattempt.md). Implémente la méthode abstraite définie dans la classe de base.

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _createAttemptResult_|Résultat d’une tentative de création.|
