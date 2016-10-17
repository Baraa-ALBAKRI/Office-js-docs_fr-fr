
# <a name="labs.core.iaction"></a>Labs.Core.IAction

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Représente une action d’atelier correspondant à une interaction entre l’utilisateur et un atelier spécifié.

```
interface IAction
```


## <a name="properties"></a>Propriétés


|||
|:-----|:-----|
| `type: string`|Type d’action effectué par l’utilisateur.|
| `options: Core.IActionOptions`|Options [Labs.Core.IActionOptions](../../reference/office-mix/labs.core.iactionoptions.md) envoyées avec l’action effectuée par l’utilisateur.|
| `result: Core.IActionResult`|Résultat [Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md) de l’action.|
| `time: number`|Heure à laquelle l’action s’est terminée, exprimée en millisecondes écoulées depuis le 01 janvier 1970 00:00:00 UTC.|
