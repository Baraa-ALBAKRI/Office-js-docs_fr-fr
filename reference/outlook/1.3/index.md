# <a name="outlook-add-in-api-requirement-set-13"></a>Ensemble de conditions requises de l’API du complément Outlook 1.3

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> **Remarque** : dans cette documentation, l’[ensemble de conditions requises](../tutorial-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente. 

## <a name="whats-new-in-13"></a>Nouveautés de la version 1.3

L’ensemble de conditions requises de la version 1.3 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.2](../1.2/index.md). Les fonctionnalités suivantes ont été ajoutées :

- Prise en charge des [commandes de complément](../../../docs/outlook/add-in-commands-for-outlook.md).
- Possibilité d’enregistrer ou de fermer un élément en cours de composition.
- Amélioration de l’objet [Body](Body.md) pour autoriser les compléments à obtenir ou à définir la totalité du corps du message.
- Nouvelles méthodes de conversion pour convertir les ID aux formats EWS et REST.
- Possibilité d’ajouter des messages de notification dans la barre d’informations sur les éléments.

### <a name="change-log"></a>Journal des modifications

- Ajout de la méthode [Body.getAsync](Body.md#getAsync) : Renvoie le corps actif dans un format spécifié.
- Ajout de la méthode [Body.setAsync](Body.md#setAsync) : Remplace l’ensemble du corps avec le texte spécifié.
- Ajout de la propriété [Office.context.officeTheme](Office.context.md#officeTheme) : permet d’accéder aux couleurs du thème Office.
- Ajout de l’objet [Event](Event.md) : transmis comme paramètre aux fonctions de commande sans IU dans un complément Outlook. Utilisé pour signaler la fin du traitement de l’événement.
- Ajout de la méthode [Office.context.mailbox.item.close](Office.context.mailbox.item.md#close) : Ferme l’élément en cours qui est composé.
- Ajout de la méthode [Office.context.mailbox.item.saveAsync](Office.context.mailbox.item.md#saveAsync) : Enregistre un élément de manière asynchrone.
- Ajout de l’objet [Office.context.mailbox.item.notificationMessages](Office.context.mailbox.item.md#notificationMessages) : Obtient les messages de notification pour un élément.
- Ajout de la méthode [Office.context.mailbox.convertToEwsId](Office.context.mailbox.md#convertToEwsId) : Convertit un ID d’élément mis en forme pour REST au format EWS.
- Ajout de la méthode [Office.context.mailbox.convertToRestId](Office.context.mailbox.md#convertToRestId) : Convertit un ID d’élément mis en forme pour EWS au format REST.
- Ajout de l’énumération [Office.MailboxEnums.ItemNotificationMessageType](Office.MailboxEnums.md#ItemNotificationMessageType) : Spécifie le type de message de notification pour un rendez-vous ou un message.
- Ajout de l’énumération [Office.MailboxEnums.RestVersion](Office.MailboxEnums.md#RestVersion) : Spécifie la version de l’API REST qui correspond à un ID d’élément au format REST.
- Ajout de l’objet [NotificationMessages](NotificationMessages.md) : fournit des méthodes pour accéder aux messages de notification dans un complément Outlook.
- Ajout du type [NotificationMessageDetails](simple-types.md#NotificationMessageDetails) : renvoyé par la méthode `NotificationMessages.getAllAsync`.

## <a name="additional-resources"></a>Ressources supplémentaires

- [Compléments Outlook](../../../docs/outlook/outlook-add-ins.md)
- [Exemples de code pour les compléments Outlook](https://dev.outlook.com/MailAppsGettingStarted/Samples)
- [Prise en main](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
