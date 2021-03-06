
# <a name="labs.components.choicecomponentattempt"></a>Labs.Components.ChoiceComponentAttempt

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Représente une tentative d’un composant de choix.

```
class ChoiceComponentAttempt extends Components.ComponentAttempt
```


## <a name="methods"></a>Méthodes




### <a name="constructor"></a>constructeur

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Crée une instance de la classe **ChoiceComponentAttempt**.

 **Paramètres**


|**Nom**|**Description**|
|:-----|:-----|
| _labs_|Instance [Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx) à utiliser avec la tentative.|
| _attemptId_|ID associé à la tentative.|
| _values_|Valeurs associées à la tentative.|

### <a name="timeout"></a>timeout

 `public function timeout(callback: Labs.Core.ILabCallback<void>): void`

Indique que l’atelier a expiré.

 **Paramètres**


|**Nom**|**Description**|
|:-----|:-----|
| _callback_|Fonctions de rappel qui se déclenchent quand le serveur reçoit le message d’expiration.|

### <a name="getsubmissions"></a>getSubmissions

 `public function getSubmissions(): Components.ChoiceComponentSubmission[]`

Récupère tous les envois précédemment effectués pour une tentative donnée.


### <a name="submit"></a>submit

 `public function submit(answer: Components.ChoiceComponentAnswer, result: Components.ChoiceComponentResult, callback: Labs.Core.ILabCallback<Components.ChoiceComponentSubmission>): void`

Envoie une nouvelle réponse notée par l’atelier. N’utilise pas l’hôte pour calculer une note.

 **Paramètres**


|**Nom**|**Description**|
|:-----|:-----|
| _answer_|Réponse pour la tentative.|
| _result_|Résultat de l’envoi.|
| _callback_|Fonction de rappel qui se déclenche une fois l’envoi reçu.|

### <a name="processaction"></a>processAction

 `public function processAction(action: Labs.Core.IAction): void`

Lance le traitement de l’action [Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md).

