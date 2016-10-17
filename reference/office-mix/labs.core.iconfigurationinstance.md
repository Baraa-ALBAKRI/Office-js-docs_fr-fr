
# <a name="labs.core.iconfigurationinstance"></a>Labs.Core.IConfigurationInstance

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Classe de base pour les instances d’une configuration de laboratoire. Une instance est l’instanciation d’une configuration d’un utilisateur donné et contient un affichage traduit de la configuration d’une exécution spécifique de l’atelier. Cet affichage peut exclure des informations masquées (conseils, réponses, etc.) et contient les ID des différentes instances.

```
interface IConfigurationInstance extends Core.IUserData
```


## <a name="properties"></a>Propriétés


|||
|:-----|:-----|
| `appVersion: Core.IVersion`|Version de l’atelier associé à cette configuration.|
| `components: Core.IComponentInstance[]`|Composants associés à l’atelier.|
| `name: string`|Nom de l’atelier.|
| `timeline: Core.ITimelineConfiguration`|Configuration de la chronologie de l’atelier.|
