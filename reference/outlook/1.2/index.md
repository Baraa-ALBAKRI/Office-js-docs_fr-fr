# <a name="outlook-add-in-api-requirement-set-12"></a>Ensemble de conditions requises de l’API du complément Outlook 1.2

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> **Remarque** : dans cette documentation, l’[ensemble de conditions requises](../tutorial-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente. 

## <a name="whats-new-in-12"></a>Nouveautés de la version 1.2

L’ensemble de conditions requises de la version 1.2 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.1](../1.1/index.md). Désormais, les compléments peuvent insérer du texte au niveau du curseur de l’utilisateur, soit dans l’objet ou le corps du message.

### <a name="change-log"></a>Journal des modifications

- Ajout de la méthode [Office.context.mailbox.item.setSelectedDataAsync](Office.context.mailbox.item.md#setSelectedDataAsync) : Insère les données dans le corps ou l’objet d’un message de manière asynchrone.
- Modification de la fonction [Office.context.mailbox.item.displayReplyAllForm](Office.context.mailbox.item.md#displayReplyAllForm) : Ajout de la propriété `attachments` dans le paramètre `formData`.
- Modification de la fonction [Office.context.mailbox.item.displayReplyForm](Office.context.mailbox.item.md#displayReplyForm) : Ajout de la propriété `attachments` dans le paramètre `formData`.

## <a name="additional-resources"></a>Ressources supplémentaires

- [Compléments Outlook](../../../docs/outlook/outlook-add-ins.md)
- [Exemples de code pour les compléments Outlook](https://dev.outlook.com/MailAppsGettingStarted/Samples)
- [Prise en main](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
