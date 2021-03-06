
# <a name="projectdocument.taskselectionchanged-event"></a>ÉvénementProjectDocument.TaskSelectionChanged
Se produit quand la sélection de tâche change dans le projet actif.

|||
|:-----|:-----|
|**Hôtes :**|Project|
|**Disponible dans l’[ensemble de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Ajouté dans**|1.0|

```js
Office.EventType.TaskSelectionChanged
```


## <a name="remarks"></a>Remarques

 **TaskSelectionChanged** est une constante d’énumération [EventType](../../reference/shared/eventtype-enumeration.md) pouvant être utilisée dans les méthodes [ProjectDocument.addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) et [ProjectDocument.removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md) pour ajouter ou supprimer un gestionnaire pour l’événement.


## <a name="example"></a>Exemple

L’exemple de code suivant ajoute un gestionnaire pour l’événement **TaskSelectionChanged**. Lorsque la sélection de tâche change dans le document, il obtient le GUID de la tâche sélectionnée.

L’exemple suppose que votre complément comporte une référence à la bibliothèque jQuery et que le contrôle de page suivant est défini dans la balise div de contenu du corps de la page.




```HTML
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            Office.context.document.addHandlerAsync(
                Office.EventType.TaskSelectionChanged,
                getTaskGuid);
            getTaskGuid();
        });
    };

    // Get the GUID of the selected task and display it in the add-in.
    function getTaskGuid() {
        Office.context.document.getSelectedTaskAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#message').html(result.value);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```

Pour obtenir un exemple qui montre comment utiliser un gestionnaire d’événements **TaskSelectionChanged** dans un complément Projet, voir l’article expliquant comment [créer votre premier complément du volet Office pour Project 2013 à l’aide d’un éditeur de texte](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## <a name="support-details"></a>Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cet événement est pris en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cet événement.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||||
|:-----|:-----|:-----|
||Office pour Bureau Windows|Office Online (dans un navigateur)|
|**Project**|v||

|||
|:-----|:-----|
|**Disponible dans les ensembles de conditions requises**|Selection|
|**Niveau d’autorisation minimal**|[Restreint](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## <a name="support-history"></a>Historique de prise en charge



|**Version**|**Modifications**|
|:-----|:-----|
|1.0|<ul><li>Introduit</li></ul>|

## <a name="see-also"></a>Voir aussi



#### <a name="other-resources"></a>Autres ressources


[Création de votre premier complément du volet Office pour Project 2013 à l’aide d’un éditeur de texte](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
[Énumération EventType](../../reference/shared/eventtype-enumeration.md)
[Méthode ProjectDocument.addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md)
[Méthode ProjectDocument.removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md)
[Objet ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
