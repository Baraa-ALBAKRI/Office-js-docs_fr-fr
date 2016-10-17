
# <a name="labs.components.iinputcomponent"></a>Labs.Components.IInputComponent

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Permet des interactions avec un composant de saisie.

```
interface IInputComponent extends Labs.Core.IComponent
```


## <a name="properties"></a>Propriétés


|Nom|Description|
|:-----|:-----|
| `maxScore: number`|Note maximale autorisée pour le composant de saisie.|
| `timeLimit: number`|Délai imparti à la résolution du problème du composant de saisie.|
| `hasAnswer: boolean`|Indique **True** si le composant a une réponse.|
| `answer: any`|Réponse au problème du composant, le cas échéant.|
| `secure: boolean`|Indique **True** si le composant de saisie est sécurisé.|
