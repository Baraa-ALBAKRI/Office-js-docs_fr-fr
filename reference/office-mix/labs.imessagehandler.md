
# <a name="labs.imessagehandler"></a>Labs.IMessageHandler

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Interface permettant de définir des gestionnaires d’événements.

```
interface IMessageHandler(origin: Window, data: any, callback: Labs.Core.ILabCallback<any>): void
```


## 

 **Paramètres**


|||
|:-----|:-----|
| `origin`|Fenêtre de l’atelier qui a émis le message.|
| `data`|Contenu du message.|
| `callback`|Fonction de rappel qui se déclenche une fois le message reçu.|
