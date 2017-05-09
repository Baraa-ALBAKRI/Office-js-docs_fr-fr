# <a name="outlook-add-in-api-requirement-set-15"></a>Ensemble de conditions requises de l’API du complément Outlook 1.5

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

## <a name="whats-new-in-15"></a>Nouveautés de la version 1.5

L’ensemble de conditions requises de la version 1.5 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.4](../1.4/index.md). Les fonctionnalités suivantes ont été ajoutées :

- Prise en charge des [volets Office épinglables](../../../docs/outlook/manifests/pinnable-taskpane.md).
- Prise en charge de l’appel des [API REST](../../../docs/outlook/use-rest-api.md).
- Possibilité de marquer une pièce jointe comme élément incorporé.
- Possibilité de fermer un volet Office ou une boîte de dialogue.

### <a name="change-log"></a>Journal des modifications

- Ajout de la méthode [Office.context.mailbox.addHandlerAsync](Office.context.mailbox.md#addHandlerAsync) : ajoute un gestionnaire d’événements pour un événement pris en charge.
- Ajout de l’énumération [Office.EventType](Office.md#EventType) : spécifie l’événement associé à un gestionnaire d’événements.
- Ajout de la propriété [Office.context.mailbox.restUrl](Office.context.mailbox.md#restUrl) : obtient l’URL du point de terminaison REST de ce compte de messagerie.
- Modification de la méthode [Office.context.mailbox.getCallbackTokenAsync](Office.context.mailbox.md#getCallbackTokenAsync) : cette nouvelle version comprend une nouvelle signature (`getCallbackTokenAsync([options], callback)`). La version d’origine est toujours disponible et reste inchangée.
- Ajout de la méthode [Office.context.ui.closeContainer](Office.context.ui.md#closeContainer) : 
- Modification de la méthode [Office.context.mailbox.item.addFileAttachmentAsync](Office.context.mailbox.item.md#addFileAttachmentAsync) : nouvelle valeur du dictionnaire `options` appelée `isInline`. Elle indique qu’une image est incorporée dans le corps du message.
- Modification de la fonction [Office.context.mailbox.item.displayReplyAllForm](Office.context.mailbox.item.md#displayReplyAllForm) : nouvelle valeur du dictionnaire `formData.attachments` appelée `isInline`. Elle indique qu’une image est incorporée dans le corps du message.
- Modification de la fonction [Office.context.mailbox.item.displayReplyForm](Office.context.mailbox.item.md#displayReplyForm) : nouvelle valeur du dictionnaire `formData.attachments` appelée `isInline`. Elle indique qu’une image est incorporée dans le corps du message.

## <a name="additional-resources"></a>Ressources supplémentaires

- [Compléments Outlook](../../../docs/outlook/outlook-add-ins.md)
- [Exemples de code pour les compléments Outlook](https://dev.outlook.com/MailAppsGettingStarted/Samples)
- [Prise en main](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
