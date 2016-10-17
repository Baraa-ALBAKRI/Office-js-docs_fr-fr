
# <a name="labs.components.choicecomponentinstance"></a>Labs.Components.ChoiceComponentInstance

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Représente une instance d’un composant de choix.

```
class ChoiceComponentInstance extends Labs.ComponentInstance<Components.ChoiceComponentAttempt>
```


## <a name="properties"></a>Propriétés


|Propriété|Description|
|:-----|:-----|
| `public var component: Components.IChoiceComponentInstance`|Instance du composant [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md) sous-jacent représentée par cette classe.|

## <a name="methods"></a>Méthodes




### <a name="constructor"></a>constructeur

 `function constructor(component: Components.IChoiceComponentInstance)`

Crée une instance de la classe **ChoiceComponentInstance**.

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _component_|Objet [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md) à utiliser pour créer cette classe.|

### <a name="buildattempt"></a>buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ChoiceComponentAttempt`

Crée une instance **ChoiceComponentAttempt** et implémente la méthode abstraite définie dans la classe de base.

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _createAttemptResult_|Résultat de l’action Tentative de création.|
