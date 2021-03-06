# <a name="onenote-javascript-api-programming-overview"></a>Vue d’ensemble de la programmation de l’API JavaScript de OneNote

OneNote présente une API JavaScript pour les compléments OneNote Online. Vous pouvez créer des compléments de volet de tâches et de contenu, ainsi que des commandes de complément qui interagissent avec les objets OneNote et se connectent à des services web ou à d’autres ressources basées sur le web.

>**Remarque :** Lorsque vous créez votre complément, si vous envisagez de le [publier](../publish/publish.md) dans Office Store, assurez-vous que vous respectez les [stratégies de validation Office Store](https://msdn.microsoft.com/en-us/library/jj220035.aspx). Par exemple, pour réussir la validation, votre complément doit fonctionner sur toutes les plateformes qui prennent en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3) et la [page relative à la disponibilité des compléments Office sur les plateformes et les hôtes](https://dev.office.com/add-in-availability)).

Les compléments sont constitués de deux composants de base :

- Une **application web** comportant une page web et les fichiers CSS, JavaScript ou autres requis. Ces fichiers sont hébergés sur un serveur web ou un service d’hébergement web, tel que Microsoft Azure. Dans OneNote Online, l’application web s’affiche dans un contrôle de navigateur ou un iFrame.
    
- Un **manifeste XML** spécifiant l’URL de la page web du complément, ainsi que les conditions d’accès, les paramètres et fonctionnalités du complément. Ce fichier est stocké sur le client. Les compléments OneNote utilisent le même format de [manifeste](https://dev.office.com/docs/add-ins/overview/add-in-manifests) que les autres compléments Office.

**Complément pour Office = manifeste + page web**

![Un complément Office se compose d’un manifeste et d’une page web](../../images/onenote-add-in.png)

### <a name="using-the-javascript-api"></a>Utilisation de l’API JavaScript

Les compléments utilisent le contexte d’exécution de l’application hôte pour accéder à l’API JavaScript. L’API comporte deux couches : 

- Une **API enrichie** pour les opérations spécifiques de OneNote, accessible via l’objet **Application**.
- Une **API commune** qui est partagée entre les applications Office, accessible via l’objet **Document**.

#### <a name="accessing-the-rich-api-through-the-application-object"></a>Accès à l’API enrichie via l’objet *Application*

Utilisez l’objet **Application** pour accéder aux objets OneNote tels que **Notebook**, **Section** et **Page**. Grâce à l’API enrichie, vous pouvez exécuter des opérations par lot sur les objets proxy. Le flux de base ressemble à ceci : 

1 - Obtenez l’instance de l’application à partir du contexte.

2 - Créez un proxy qui représente l’objet OneNote que vous souhaitez utiliser. Vous interagissez simultanément avec les objets proxy en lisant et en écrivant leurs propriétés et en appelant leurs méthodes. 

3 - Appelez la méthode **load** sur le serveur proxy pour la remplir avec les valeurs de propriété spécifiées dans le paramètre. Cet appel est ajouté à la file d’attente des commandes. 

   Les appels de méthode à l’API (tels que `context.application.getActiveSection().pages;`) sont également ajoutés à la file d’attente.
    
4 - Appelez la méthode **context.sync** pour exécuter toutes les commandes en attente dans l’ordre dans lequel elles ont été mises en file d’attente. Cela permet de synchroniser l’état entre votre script d’exécution et les objets réels, en récupérant les propriétés des objets OneNote chargés à utiliser dans vos scripts. Vous pouvez utiliser l’objet Promise renvoyé pour créer une chaîne avec les actions supplémentaires.

Par exemple : 

```
    function getPagesInSection() {
        OneNote.run(function (context) {
            
            // Get the pages in the current section.
            var pages = context.application.getActiveSection().pages;
            
            // Queue a command to load the id and title for each page.            
            pages.load('id,title');
            
            // Run the queued commands, and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    
                    // Read the id and title of each page. 
                    $.each(pages.items, function(index, page) {
                        var pageId = page.id;
                        var pageTitle = page.title;
                        console.log(pageTitle + ': ' + pageId); 
                    });
                })
                .catch(function (error) {
                    app.showNotification("Error: " + error);
                    console.log("Error: " + error);
                    if (error instanceof OfficeExtension.Error) {
                        console.log("Debug info: " + JSON.stringify(error.debugInfo));
                    }
                });
        });
    }
```

Vous pouvez déterminer les objets et les opérations OneNote pris en charge dans la [référence de l’API](../../reference/onenote/onenote-add-ins-javascript-reference.md).

### <a name="accessing-the-common-api-through-the-document-object"></a>Accès à l’API commune via l’objet *Document*

Utilisez l’objet **Document** pour accéder à l’API commune, par exemple les méthodes [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) et [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync). 

Par exemple :  

```
function getSelectionFromPage() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(error.message);
            }
            else $('#input').val(asyncResult.value);
        });
}
```
Les compléments OneNote prennent en charge uniquement les API communes suivantes :

| API | Commentaires |
|:------|:------|
| [Office.context.document.getSelectedDataAsync](https://msdn.microsoft.com/en-us/library/office/fp142294.aspx) | **Office.CoercionType.Text** et **Office.CoercionType.Matrix** uniquement |
| [Office.context.document.setSelectedDataAsync](https://msdn.microsoft.com/en-us/library/office/fp142145.aspx) | **Office.CoercionType.Text**, **Office.CoercionType.Image** et **Office.CoercionType.Html** uniquement | 
| [var mySetting = Office.context.document.settings.get(name);](https://msdn.microsoft.com/en-us/library/office/fp142180.aspx) | Les paramètres sont pris en charge par les compléments de contenu uniquement | 
| [Office.context.document.settings.set(name, value);](https://msdn.microsoft.com/en-us/library/office/fp161063.aspx) | Les paramètres sont pris en charge par les compléments de contenu uniquement | 
| [Office.EventType.DocumentSelectionChanged](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) ||En règle générale, vous utilisez uniquement l’API commune pour effectuer une action qui n’est pas prise en charge dans l’API enrichie. Pour en savoir plus sur l’utilisation de l’API commune, voir la [documentation](https://dev.office.com/docs/add-ins/overview/office-add-ins) et les [références](https://dev.office.com/reference/add-ins/javascript-api-for-office) concernant les compléments Office.


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a>Diagramme du modèle objet OneNote 
Le diagramme suivant représente ce qui est actuellement disponible dans l’API JavaScript de OneNote.

  ![Diagramme du modèle objet OneNote](../../images/onenote-om.png)


## <a name="additional-resources"></a>Ressources supplémentaires

- [Créer votre premier complément OneNote](onenote-add-ins-getting-started.md)
- [Référence de l’API JavaScript de OneNote](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Exemple de grille d’évaluation](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Vue d’ensemble de la plateforme des compléments Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
