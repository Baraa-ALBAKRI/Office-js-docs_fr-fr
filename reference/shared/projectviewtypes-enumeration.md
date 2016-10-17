
# <a name="projectviewtypes-enumeration"></a>ProjectViewTypes, énumération
Spécifie les types d’affichage que la méthode **[getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)** peut reconnaître.

|||
|:-----|:-----|
|**Hôtes :**|Project|
|**Ajouté dans**|1.0|

```
ProjectViewTypes={
    Gantt           : 1, 
    NetworkDiagram  : 2, 
    TaskDiagram     : 3, 
    TaskForm        : 4, 
    TaskSheet       : 5, 
    ResourceForm    : 6, 
    ResourceSheet   : 7, 
    ResourceGraph   : 8, 
    TeamPlanner     : 9, 
    TaskDetails     : 10, 
    TaskNameForm    : 11, 
    ResourceNames   : 12, 
    Calendar        : 13, 
    TaskUsage       : 14, 
    ResourceUsage   : 15, 
    Timeline        : 16
}
```


## <a name="members"></a>Membres


****


|**Membre**|**Description**|
|:-----|:-----|
|**Gantt**|Affichage Diagramme de Gantt.|
|**NetworkDiagram**|Affichage Réseau de tâches.|
|**TaskDiagram**|Affichage Diagramme de tâches.|
|**TaskForm**|Affichage Formulaire de tâche.|
|**TaskSheet**|Affichage Tableau des tâches.|
|**ResourceForm**|Affichage Formulaire ressource.|
|**ResourceSheet**|Affichage Tableau des ressources.|
|**ResourceForm**|Affichage Formulaire ressource.|
|**ResourceGraph**|Affichage Graphe des ressources.|
|**TeamPlanner**|Affichage Planificateur d’équipe.|
|**TaskDetails**|Affichage Détails des tâches.|
|**TaskNameForm**|Affichage Fiche nom de tâche.|
|**ResourceNames**|Affichage Noms des ressources.|
|**Calendar**|Affichage Calendrier|
|**TaskUsage**|Affichage Utilisation des tâches|
|**ResourceUsage**|Affichage Utilisation des ressources.|
|**Timeline**|Affichage Chronologie.|

## <a name="remarks"></a>Remarques

La méthode **[getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)** renvoie le nom et la valeur de constante de **ProjectViewTypes** correspondant à l’affichage actif.


## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette énumération est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette énumération.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour bureau Windows**|**Office Online (dans un navigateur)**|
|:-----|:-----|:-----|
|**Project**|v||

|||
|:-----|:-----|
|**Types de complément**|Volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## <a name="support-history"></a>Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.0|Introduit|

## <a name="see-also"></a>Voir aussi



#### <a name="other-resources"></a>Autres ressources


[Méthode getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)
