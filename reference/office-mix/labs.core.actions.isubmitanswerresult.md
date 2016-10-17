
# <a name="labs.core.actions.isubmitanswerresult"></a>Labs.Core.Actions.ISubmitAnswerResult

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Résultat de l’envoi d’une réponse pour une tentative.

```
interface ISubmitAnswerResult extends Core.IActionResult
```


## <a name="properties"></a>Propriétés


|||
|:-----|:-----|
| `submissionId: string`|ID associé à l’envoi. Fourni par le serveur.|
| `complete: boolean`|Renvoie  **true** si la tentative est terminée grâce à l’envoi en cours.|
| `score: any`|Informations sur la note associée à l’envoi.|
