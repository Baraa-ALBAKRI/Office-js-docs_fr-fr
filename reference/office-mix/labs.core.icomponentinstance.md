
# <a name="labs.core.icomponentinstance"></a>Labs.Core.IComponentInstance

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Classe de base pour les instances des composants de l’atelier.

```
interface IComponentInstance extends Core.ILabObject, Core.IUserData
```


## <a name="properties"></a>Propriétés


|||
|:-----|:-----|
| `componentId: string`|ID du composant associé à cette instance.|
| `name: string`|Nom du composant.|
| `values: {[type:string]: Core.IValueInstance[]}`|Mappage des propriétés de valeur associées au composant.|

## <a name="remarks"></a>Remarques

Une instance de composant est l’instanciation d’un composant d’un utilisateur. Il contient un affichage traduit du composant pour une exécution spécifique de l’atelier. Cet affichage peut exclure des informations masquées (réponses, conseils, etc.) et contient les ID de différentes instances.

