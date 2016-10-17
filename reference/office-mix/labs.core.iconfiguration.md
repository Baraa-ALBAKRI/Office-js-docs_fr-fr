
# <a name="labs.core.iconfiguration"></a>Labs.Core.IConfiguration

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Structure de données de configuration de laboratoire.

```
interface IConfiguration extends Core.IUserData
```


## <a name="properties"></a>Propriétés


|||
|:-----|:-----|
| `appVersion: Core.IVersion`|Version de l’application associée à cette configuration.|
| `components: Core.IComponent[]`|Composants inclus avec l’atelier.|
| `name: string`|Nom de l’atelier.|
| `timeline: Core.ITimelineConfiguration`|Configuration de la chronologie de l’atelier.|
| `analytics: Core.IAnalyticsConfiguration`|Configuration des analyses de l’atelier.|
