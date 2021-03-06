
# <a name="officetheme.bodybackgroundcolor-property"></a>Propriété officeTheme.bodyBackgroundColor
Obtient la couleur d’arrière-plan du corps du thème Office.

 **Important :** Cette API fonctionne actuellement uniquement dans Excel, Outlook, PowerPoint et Word dans [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) sur le bureau Windows.


|||
|:-----|:-----|
|**Hôtes :**|Excel, Outlook, PowerPoint, Word|
|**Disponible dans l’[ensemble de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Pas dans un ensemble|
|**Ajouté dans**|1.3|



```
var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
```


## <a name="return-value"></a>Valeur renvoyée

Triplet de couleur hexadécimal.


## <a name="remarks"></a>Remarques

Les couleurs renvoyées correspondent aux valeurs du thème Office, sélectionné par l’utilisateur en accédant à **Fichier**  >  **Compte Office**  >  **Thème Office**, qui est appliqué à toutes les applications hôtes Office.


## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|**OWA pour les appareils**|
|:-----|:-----|:-----|:-----|:-----|
|**Excel**|v||||
|**Outlook**|v||||
|**PowerPoint**|v||||
|**Word**|v||||

|||
|:-----|:-----|
|**Niveau d’autorisation minimal**|[Restreint](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|De contenu, de volet de tâche, Outlook|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## <a name="support-history"></a>Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.3|Introduit|
