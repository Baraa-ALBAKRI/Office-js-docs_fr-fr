# <a name="outlook-add-in-api-requirement-set-11"></a>Ensemble de conditions requises de l’API du complément Outlook 1.1

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> **Remarque** : dans cette documentation, l’[ensemble de conditions requises](../tutorial-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente. 

## <a name="whats-new-in-11"></a>Nouveautés de la version 1.1

L’ensemble de conditions requises de la version 1.1 comprend toutes les fonctionnalités de l’ensemble de conditions requises de la version 1.0. Désormais, les compléments peuvent accéder au corps des messages et des rendez-vous et vous pouvez modifier l’élément actif.

### <a name="change-log"></a>Journal des modifications

- Ajout de l’objet [Body](Body.md) : fournit des méthodes pour ajouter et mettre à jour le contenu d’un élément dans un complément Outlook.
- Ajout de l’objet [Location](Location.md) : Fournit des méthodes pour obtenir et définir le lieu d’une réunion dans un complément Outlook.
- Ajout de l’objet [Recipients](Recipients.md) : fournit des méthodes pour obtenir et définir les destinataires d’un rendez-vous ou d’un message dans un complément Outlook.
- Ajout de l’objet [Subject](Subject.md) : Fournit des méthodes pour obtenir et définir l’objet d’un rendez-vous ou d’un message dans un complément Outlook.
- Ajout de l’objet [Time](Time.md) : fournit des méthodes pour obtenir et définir l’heure de début ou de fin d’une réunion dans un complément Outlook.
- Ajout de la méthode [Office.context.mailbox.item.addFileAttachmentAsync](Office.context.mailbox.item.md#addFileAttachmentAsync) : ajoute un fichier à un message ou un rendez-vous en pièce jointe.
- Ajout de la méthode [Office.context.mailbox.item.addItemAttachmentAsync](Office.context.mailbox.item.md#addItemAttachmentAsync) : ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.
- Ajout de la méthode [Office.context.mailbox.item.removeAttachmentAsync](Office.context.mailbox.item.md#removeAttachmentAsync) : supprime une pièce jointe d’un message ou d’un rendez-vous.
- Ajout de l’objet [Office.context.mailbox.item.body](Office.context.mailbox.item.md#body) : obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.
- Ajout de l’objet [Office.context.mailbox.item.bcc](Office.context.mailbox.item.md#bcc) : obtient ou définit les destinataires en Cci (copie carbone invisible) d’un message.
- Ajout de l’énumération [Office.MailboxEnums.RecipientType](Office.MailboxEnums.md#RecipientType) : spécifie le type de destinataire d’un rendez-vous.

## <a name="additional-resources"></a>Ressources supplémentaires

- [Compléments Outlook](../../../docs/outlook/outlook-add-ins.md)
- [Exemples de code pour les compléments Outlook](https://dev.outlook.com/MailAppsGettingStarted/Samples)
- [Prise en main](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
