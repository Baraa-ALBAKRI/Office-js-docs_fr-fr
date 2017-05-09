# <a name="outlook-add-in-api-requirement-set-14"></a>Ensemble de conditions requises de l’API du complément Outlook 1.4

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> **Remarque** : dans cette documentation, l’[ensemble de conditions requises](../tutorial-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente.

## <a name="whats-new-in-14"></a>Nouveautés de la version 1.4

L’ensemble de conditions requises de la version 1.4 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.3](../1.3/index.md). Il comprend en plus l’accès à l’espace de noms `Office.ui`.

### <a name="change-log"></a>Journal des modifications

- Ajout de la méthode [Office.context.ui.displayDialogAsync](../../shared/officeui.displaydialogasync.md) : Affiche une boîte de dialogue dans un hôte Office.
- Ajout de la méthode [Office.context.ui.messageParent](../../shared/officeui.messageparent.md) : Remet un message de la part de la boîte de dialogue à sa page parent/d’ouverture.
- Ajout de l’objet [Dialog](../../shared/officeui.dialog.md) : objet renvoyé lorsque la méthode `displayDialogAsync` est appelée.

## <a name="additional-resources"></a>Ressources supplémentaires

- [Compléments Outlook](../../../docs/outlook/outlook-add-ins.md)
- [Exemples de code pour les compléments Outlook](https://dev.outlook.com/MailAppsGettingStarted/Samples)
- [Prise en main](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
