
# <a name="documentactiveviewchanged-event"></a>Événement Document.ActiveViewChanged
Survient lorsque l’utilisateur modifie l’affichage actuel du document.

|||
|:-----|:-----|
|**Hôtes :**|PowerPoint|
|**Nouveauté de**|1.1|

```
Office.EventType.ActiveViewChanged
```


## <a name="remarks"></a>Remarques

Pour ajouter un gestionnaire d’événements à l’événement **ActiveViewChanged** d’un document, utilisez la méthode [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) de l’objet **Document**. Le gestionnaire d’événements reçoit un argument de type [ActiveViewChangedEventArgs](../../reference/shared/document.activeviewchangedeventargs.md).


## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour bureau Windows**|**Office Online (dans un navigateur)**|**Office pour Mac**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|:-----|
|**PowerPoint**|v||v|Y|

>**Remarque : Cet événement ne se déclenche pas dans les scénarios PowerPoint Online, car le mode diaporama est considéré comme une nouvelle session. Pour obtenir la vue active, vous devez envoyer la requête correspondante pendant Office.Initialize.
 

|||
|:-----|:-----|
|**Nouveauté de**|1.1|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|
