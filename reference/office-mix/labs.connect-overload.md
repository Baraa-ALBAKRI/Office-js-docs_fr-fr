
# <a name="labs.connect-(overload)"></a>Labs.Connect (surcharge)

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Initialise une connexion avec l’hôte.

```
function connect(labHost: Core.ILabHost, callback: Core.ILabCallback<Core.IConnectionResponse>)
```


## <a name="parameters"></a>Paramètres


|||
|:-----|:-----|
| _labHost_|Facultatif. Instance [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md) à laquelle se connecter. Si l’hôte n’est pas spécifié, un hôte est créé avec [Labs.DefaultHostBuilder](../../reference/office-mix/labs.defaulthostbuilder.md).|
| _callback_|Rappel qui se déclenche une fois la connexion établie.|

## <a name="return-value"></a>Valeur renvoyée

Renvoie une connexion à l’hôte.

