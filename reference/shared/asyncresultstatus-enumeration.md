
# <a name="asyncresultstatus-enumeration"></a>AsyncResultStatus, énumération
Spécifie le résultat d’un appel asynchrone. 

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Dernière modification dans**|1.1|

```
Office.AsyncResultStatus
```


## <a name="members"></a>Membres


**Valeurs**


|**Énumération**|**Valeur**|**Description**|
|:-----|:-----|:-----|
|Office.AsyncResultStatus.Succeeded|"succeeded"|L’appel a réussi.|
|Office.AsyncResultStatus.Failed|"failed"|L’appel n’a pas réussi.|

## <a name="remarks"></a>Remarques

Retourné par la propriété [status](../../reference/shared/asyncresult.status.md) de l’objet [AsyncResult](../../reference/shared/asyncresult.md).


## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette énumération est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette énumération.


Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|**OWA pour les appareils**|**Office pour Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**|v|||||
|**Excel**|v|v|v|||
|**Outlook**|v|v||v|v|
|**PowerPoint**|v|v|v|||
|**Project**|v|||||
|**Word**|v||v|||

|||
|:-----|:-----|
|**Types de complément**|De contenu, de volet de tâche, Outlook|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## <a name="support-history"></a>Historique de prise en charge


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire des compléments pour Access.|
|1.0|Introduit|
