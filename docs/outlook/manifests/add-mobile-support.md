# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>Ajouter la prise en charge des commandes de complément pour Outlook Mobile

> **Remarque** : les commandes de complément pour Outlook Mobile sont actuellement uniquement prises en charge par Outlook pour iOS.

Grâce aux commandes de complément dans Outlook Mobile, vos utilisateurs peuvent accéder aux mêmes fonctionnalités (avec certaines [limitations](#code-considerations)) que celles dont ils disposent déjà dans Outlook pour Windows, Outlook pour Mac et Outlook sur le web. L’ajout de la prise en charge d’Outlook Mobile nécessite la mise à jour du manifeste de complément et éventuellement la modification de votre code pour les scénarios mobiles.

## <a name="updating-the-manifest"></a>Mise à jour du manifeste

La première étape de l’activation des commandes de complément dans Outlook Mobile est de les définir dans le manifeste du complément. Le schéma **VersionOverrides** v1.1 définit un nouveau facteur de forme pour les versions mobiles, [MobileFormFactor](../../reference/manifest/mobileformfactor.md).

Cet élément contient toutes les informations pour charger le complément dans des clients mobiles. Cela vous permet de définir entièrement différents éléments de l’interface utilisateur et fichiers JavaScript pour l’expérience mobile.

L’exemple suivant montre un bouton d’un volet des tâches unique dans un élément **MobileFormFactor**.

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
  ...
  <MobileFormFactor>
    <FunctionFile resid="residUILessFunctionFileUrl" />
    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
      <Control xsi:type="MobileButton" id="TaskPane1Btn">
        <Label resid="residTaskPaneButton0Name" />
        <Icon xsi:type="bt:MobileIconList">
          <bt:Image size="25" scale="1" resid="tp0icon" />
          <bt:Image size="25" scale="2" resid="tp0icon" />
          <bt:Image size="25" scale="3" resid="tp0icon" />

          <bt:Image size="32" scale="1" resid="tp0icon" />
          <bt:Image size="32" scale="2" resid="tp0icon" />
          <bt:Image size="32" scale="3" resid="tp0icon" />

          <bt:Image size="48" scale="1" resid="tp0icon" />
          <bt:Image size="48" scale="2" resid="tp0icon" />
          <bt:Image size="48" scale="3" resid="tp0icon" />
        </Icon>
        <Action xsi:type="ShowTaskpane">
          <SourceLocation resid="residTaskpaneUrl" />
        </Action>
      </Control>
    </ExtensionPoint>
  </MobileFormFactor>
  ...
</VersionOverrides>
```

Cet exemple est semblable aux éléments qui apparaissent dans un élément [DesktopFormFactor](../../reference/manifest/desktopformfactor.md), avec toutefois quelques différences importantes.

- L’élément [OfficeTab](../../reference/manifest/officetab.md) n’est pas utilisé.
- L’élément [ExtensionPoint](../../reference/manifest/exensionpoint.md) doit avoir un seul élément enfant. Si le complément ajoute uniquement un bouton, l’élément enfant doit être un élément [Control](../../reference/manifest/control.md). Si le complément ajoute plusieurs boutons, l’élément enfant doit être un élément [Group](../../reference/manifest/group.md) qui contient plusieurs éléments `Control`.
- Il n’existe aucun équivalent de type `Menu` pour l’élément `Control`.
- L’élément [Supertip](../../reference/manifest/supertip.md) n’est pas utilisé.
- Les tailles d’icône requises sont différentes. Au minimum, les compléments mobiles doivent prendre en charge les icônes 25 x 25, 32 x 32 et 48 x 48 pixels.

## <a name="code-considerations"></a>Éléments à prendre en compte pour le code

La conception d’un complément pour mobile implique certaines considérations supplémentaires.

### <a name="use-rest-instead-of-exchange-web-services"></a>Utiliser REST plutôt que les services web Exchange

La méthode [Office.context.mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) n’est pas prise en charge dans Outlook Mobile. Les compléments doivent privilégier l’obtention d’informations auprès de l’API Office.js lorsque cela est possible. Si les compléments requièrent des informations non exposées par l’API Office.js, ils doivent utiliser les [API REST Outlook](https://dev.outlook.com/restapi/reference) pour accéder à la boîte aux lettres de l’utilisateur. 

L’ensemble de conditions de boîte aux lettres 1.5 introduit une nouvelle version d’[Office.context.mailbox.getCallbackTokenAsync](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#getCallbackTokenAsync) qui peut demander un jeton d’accès compatible avec les API REST et une nouvelle propriété [Office.context.mailbox.restUrl](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#restUrl) qui peut être utilisée pour rechercher le point de terminaison de l’API REST pour l’utilisateur.

### <a name="pinch-zoom"></a>Pincer pour zoomer

Par défaut les utilisateurs peuvent utiliser le mouvement pincer pour zoomer sur les volets des tâches. Si ce mouvement n’est pas pertinent pour votre scénario, veillez à désactiver la fonction « pincer pour zoomer » dans votre code HTML.

### <a name="closing-taskpanes"></a>Fermeture des volets des tâches

Dans Outlook Mobile, les volets des tâches occupent la totalité de l’écran et exigent par défaut que l’utilisateur les ferme pour revenir au message. Envisagez d’utiliser la méthode [Office.context.ui.closeContainer](https://dev.outlook.com/reference/add-ins/1.5/Office.context.ui.html#closeContainer) pour fermer le volet des tâches lorsque votre scénario est terminée.

### <a name="compose-mode-and-appointments"></a>Mode composition et rendez-vous

Actuellement, les compléments dans Outlook Mobile ne prennent en charge l’activation que lors de la lecture des messages. Les compléments ne sont pas activés lors de la composition des messages, ou lors de l’affichage ou de la rédaction des rendez-vous.

### <a name="unsupported-apis"></a>API non prises en charge

Les API suivantes ne sont pas prises en charge par Outlook Mobile.

  - [Office.context.officeTheme](../../reference/outlook/Office.context.md)
  - [Office.context.mailbox.ewsUrl](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.convertToEwsId](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.convertToRestId](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.displayAppointmentForm](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.displayMessageForm](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.item.dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.resources](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.displayReplyAllForm](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.displayReplyForm](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getEntities](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getEntitiesByType](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getFilteredEntitiesByName](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getRegexMatches](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getRegexMatchesByName](../../reference/outlook/Office.context.mailbox.item.md)