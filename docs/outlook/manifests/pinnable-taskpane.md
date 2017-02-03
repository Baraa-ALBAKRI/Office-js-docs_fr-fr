# <a name="implement-a-pinnable-taskpane-in-outlook"></a>Mettre en œuvre un volet des tâches épinglable dans Outlook

La commande de forme UX [taskpane](../add-in-commands-for-outlook.md#launching-a-task-pane) pour complément ouvre un volet des tâches vertical à droite d’un message ou rendez-vous ouvert, ce qui permet au complément de fournir une interface utilisateur pour des interactions plus détaillées (remplissage de plusieurs champs, etc.). Ce volet des tâches peut être affiché dans le volet de lecture lorsque vous affichez une liste des messages, ce qui permet un traitement rapide d’un message.

Toutefois, par défaut, si un utilisateur a un complément de volet des tâches ouvert pour un message dans le volet de lecture et sélectionne un nouveau message, le volet des tâches est automatiquement fermé. Pour un complément très sollicité, l’utilisateur peut préférer conserver ce volet ouvert, supprimant ainsi le besoin de réactiver le complément sur chaque message. Avec les volets des tâches épinglables, votre complément peut donner à l’utilisateur cette option.

> **Remarque** : les volets de tâches épinglables ne sont actuellement pris en charge que par Outlook 2016 pour Windows (build 7628.1000 ou version ultérieure).

## <a name="support-taskpane-pinning"></a>Prise en charge de l’épinglage des volets des tâches

La première étape consiste à ajouter une prise en charge de l’épinglage, ce qui est effectué dans le [manifeste](./manifests.md) du complément. Cette opération est effectuée en ajoutant l’élément [SupportsPinning](../../../reference/manifest/action.md#supportspinning) à l’élément `Action` qui décrit le bouton du volet des tâches.

L’élément `SupportsPinning` est défini dans le schéma VersionOverrides V1.1, vous devez donc inclure un élément [VersionOverrides](../../../reference/manifest/versionoverrides.md) à la fois pour les versions 1.0 et 1.1.

> **Remarque :** Si vous envisagez de [publier](../../publish/publish.md) votre complément Outlook sur l’Office Store, lorsque vous utilisez l’élément **SupportsPinning** afin d’obtenir la [validation de l’Office Store](https://msdn.microsoft.com/en-us/library/jj220035.aspx), le contenu de votre complément ne doit pas être statique et doit afficher clairement les données liées au message qui est ouvert ou sélectionné dans la boîte aux lettres.

```xml
<!-- Task pane button -->
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
    <SupportsPinning>true</SupportsPinning>
  </Action>
</Control>
```

Pour obtenir un exemple complet, consultez le contrôle `msgReadOpenPaneButton` dans l’[exemple de manifeste command-demo](https://github.com/jasonjoh/command-demo/blob/master/command-demo-manifest.xml).

## <a name="handling-ui-updates-based-on-currently-selected-message"></a>Gestion des mises à jour de l’interface utilisateur en fonction du message actuellement sélectionné

Pour mettre à jour l’interface utilisateur ou les variables internes de votre volet des tâches en fonction de l’élément actif, vous devez enregistrer un gestionnaire d’événements pour être notifié de la modification.

### <a name="implement-the-event-handler"></a>Mettre en œuvre le gestionnaire d’événements

Le gestionnaire d’événements doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` de cet objet est réglée sur `Office.EventType.ItemChanged`. Lorsque l’événement est appelé, l’objet `Office.context.mailbox.item` est déjà mis à jour pour refléter l’élément actuellement sélectionné.

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

### <a name="register-the-event-handler"></a>Enregistrer le gestionnaire d’événements

Utilisez la méthode [Office.context.mailbox.addHandlerAsync](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#addHandlerAsync) pour inscrire votre gestionnaire d’événements pour l’événement `Office.EventType.ItemChanged`. Cette opération doit être effectuée dans la fonction `Office.initialize` de votre volet des tâches.

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## <a name="additional-resources"></a>Ressources supplémentaires

Pour un exemple de complément qui implémente un volet des tâches épinglable, consultez [command-demo](https://github.com/jasonjoh/command-demo) sur GitHub.