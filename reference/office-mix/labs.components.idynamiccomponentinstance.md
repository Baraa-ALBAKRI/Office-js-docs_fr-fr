
# <a name="labs.components.idynamiccomponentinstance"></a>Labs.Components.IDynamicComponentInstance

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Instance d’un composant dynamique.

```
interface IDynamicComponentInstance extends Labs.Core.IComponentInstance
```


## <a name="properties"></a>Propriétés


|Nom|Description|
|:-----|:-----|
| `generatedComponentTypes: string[]`|Tableau qui contient les types de composants pouvant être générés par ce composant dynamique.|
| `maxComponents: number`|Nombre maximal de composants généré par ce composant dynamique. Ou  **Labs.Components.Infinite** s’il n’y a pas de plafond.|
