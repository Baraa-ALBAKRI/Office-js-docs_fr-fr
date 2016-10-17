
# <a name="labs.core.icomponent"></a>Labs.Core.IComponent

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Classe de base pour la représentation des composants d’un laboratoire.

```
interface IComponent extends Core.ILabObject, Core.IUserData
```


## <a name="properties"></a>Propriétés


|||
|:-----|:-----|
| `name: string`|Nom du composant.|
| `values: {[type:string]: Core.IValue[]}`|Mappage des propriétés de valeur associées au composant.|
