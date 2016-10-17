
# <a name="labs.components.idynamiccomponent"></a>Labs.Components.IDynamicComponent

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Permet des interactions avec un composant dynamique.

```
interface IDynamicComponent extends Labs.Core.IComponent
```


## <a name="properties"></a>Propriétés


|Nom|Description|
|:-----|:-----|
| `generatedComponentTypes: string[]`|Tableau qui contient les types de composants pouvant être générés par ce composant dynamique.|
| `maxComponents: number`|Nombre maximal de composants généré par ce composant dynamique. Ou  **Labs.Components.Infinite** s’il n’y a pas de plafond.|
