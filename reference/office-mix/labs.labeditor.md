
# <a name="labs.labeditor"></a>Labs.LabEditor

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

L’objet **LabEditor** vous permet de modifier un atelier donné, ainsi que d’obtenir et de définir des données de configuration associées à l’atelier.

```
class LabEditor
```


## <a name="methods"></a>Méthodes


### <a name="getconfiguration"></a>getConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

Récupère la configuration de l’atelier en cours.

 **Paramètres**


|**Nom**|**Description**|
|:-----|:-----|
| _callback_|Fonction de rappel déclenchée une fois la configuration récupérée.|

### <a name="setconfiguration"></a>setConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

Définit une nouvelle configuration pour l’atelier.

 **Paramètres**


|**Nom**|**Description**|
|:-----|:-----|
| _configuration_|Configuration à définir.|
| _callback_|Fonction de rappel déclenchée une fois la configuration définie.|

### <a name="done"></a>fait

 `public function done(callback: Labs.Core.ILabCallback<void>): void`

Indique que l’utilisateur a fini de modifier l’atelier.

 **Paramètres**


|**Nom**|**Description**|
|:-----|:-----|
| _callback_|Fonction de rappel qui se déclenche quand l’éditeur de l’atelier a terminé.|
