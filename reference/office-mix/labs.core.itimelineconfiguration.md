
# <a name="labs.core.itimelineconfiguration"></a>Labs.Core.ITimelineConfiguration

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Options de configuration de l’instance [Labs.Timeline](../../reference/office-mix/labs.timeline.md). Permet de spécifier un ensemble d’options de configuration de chronologie.

```
interface ITimelineConfiguration
```


## <a name="properties"></a>Propriétés


|||
|:-----|:-----|
| `duration: number`|Durée de l’atelier, en secondes.|
| `capabilities: string[]`|Liste de tableaux des fonctionnalités de chronologie prises en charge par l’atelier (lire, suspendre, rechercher, etc.).|
