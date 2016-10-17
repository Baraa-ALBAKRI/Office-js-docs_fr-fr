
# <a name="labs.componentinstance"></a>Labs.ComponentInstance

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Représente l’instance d’un composant, qui est une instanciation d’un composant donné pour un utilisateur lors de l’exécution. L’objet comporte une vue traduite du composant pour une exécution spécifique de l’atelier.

```
class ComponentInstance<T> extends Labs.ComponentInstanceBase
```


## <a name="properties"></a>Propriétés

Aucun.


## <a name="methods"></a>Méthodes




### <a name="constructor"></a>Constructeur

 `function constructor()`

Initialise une nouvelle instance de la classe **ComponentInstance**.


### <a name="createattempt"></a>createAttempt

 `public function createAttempt(callback: Labs.Core.ILabCallback<T>): void`

Crée une tentative dans le contexte d’un composant.

 **Paramètres**


|**Nom**|**Description**|
|:-----|:-----|
| _callback_|Rappel déclenché lors de la création de la tentative.|

### <a name="getattempts"></a>getAttempts

 `public function getAttempts(callback: Labs.Core.ILabCallback<T[]>): void`

Récupère toutes les tentatives associées au composant donné.

 **Paramètres**


|**Nom**|**Description**|
|:-----|:-----|
| _callback_|Rappel déclenché lors de la récupération des tentatives.|

### <a name="getcreateattemptoptions"></a>getCreateAttemptOptions

 `public function getCreateAttemptOptions(): Labs.Core.Actions.ICreateAttemptOptions`

Récupère les options par défaut de la tentative de création. Peut être remplacé par des classes dérivées.


### <a name="buildattempt"></a>buildAttempt

 `public function buildAttempt(createAttemptResult: Labs.Core.IAction): T`

Génère une tentative à partir de l’action donnée. Doit être implémenté par des classes dérivées.

 **Paramètres**


|**Nom**|**Description**|
|:-----|:-----|
| _createAttemptResult_|Action Tentative de création pour la tentative spécifiée.|
