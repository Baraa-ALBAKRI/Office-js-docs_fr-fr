
# <a name="labs.command"></a>Labs.Command

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Commande générale permettant de transmettre des messages entre le client et l’hôte.

```
class Command
```


## <a name="properties"></a>Propriétés


|**Nom**|**Description**|
|:-----|:-----|
| `public var type: string`|Type de la commande.|
| `public var commandData: any`|Données facultatives associées à la commande.|

## <a name="methods"></a>Méthodes




### <a name="constructor"></a>constructeur

 `function constructor(type: string, commandData?: any)`

Description

 **Paramètres**


|||
|:-----|:-----|
| `type`|Type de la commande.|
| `commandData`|Données facultatives associées à la commande.|
