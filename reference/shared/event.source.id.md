

# <a name="event.source.id"></a>event.source.id
Obtient l’ID du contrôle qui a déclenché l’appel de cette fonction.

****

|||
|:-----|:-----|
|**Hôtes :**Outlook|**Type de complément :** Outlook|
|**Disponible dans les [ensembles de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Boîte aux lettres|
|**Dernière modification dans la boîte aux lettres**|1.3|
|**Modes Outlook applicables**|Lire et composer|



```js
event.source.id;
```


## <a name="return-value"></a>Valeur renvoyée

ID du contrôle qui a déclenché l’appel de cette fonction. L’ID est issu du manifeste.


## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans le tableau suivant indique que cette propriété est prise en charge dans l’application hôte Outlook correspondante. Une cellule vide indique que l’application hôte Outlook ne prend pas en charge cette propriété.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).

 **Important :** les commandes de complément et les API associées fonctionnent actuellement uniquement dans Outlook dans [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) sur le bureau Windows.


**Hôtes pris en charge par la plateforme**

| |**Office pour bureau Windows**|**Office Online (dans un navigateur)**|**OWA pour les appareils**|
|:-----|:-----|:-----|:-----|
|**Outlook**|v|||

|||
|:-----|:-----|
|**Disponible dans les ensembles de conditions requises**|Boîte aux lettres|
|**Niveau d’autorisation minimal**|[ReadWriteItem](../../docs/outlook/understanding-outlook-add-in-permissions.md)|
|**Types de complément**|Outlook|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## <a name="support-history"></a>Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.3|Introduit|
