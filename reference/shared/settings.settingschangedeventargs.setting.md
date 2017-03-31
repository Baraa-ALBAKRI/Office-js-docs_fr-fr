

# <a name="settingschangedeventargssettings-property"></a>Propriété SettingsChangedEventArgs.settings
Obtient un objet **Settings** qui représente les paramètres qui ont déclenché l’événement **settingsChanged**.

|||
|:-----|:-----|
|**Hôtes :**|Excel, Word|
|**Disponible dans l’[ensemble de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Paramètres|
|**Dernière modification dans**|1.0|

```js
var mySettings = eventArgsObj.settings;
```


## <a name="return-value"></a>Valeur renvoyée

Objet [Settings](../../reference/shared/document.settings.md) qui représente les paramètres qui ont déclenché l’événement [settingsChanged](../../reference/shared/settings.settingschangedevent.md).


## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette propriété est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette propriété.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office pour bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||v||
|**Word**|v|v||

|||
|:-----|:-----|
|**Disponible dans les ensembles de conditions requises**|Paramètres|
|**Niveau d’autorisation minimal**|[Restreint](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## <a name="support-history"></a>Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.0|Introduit|
