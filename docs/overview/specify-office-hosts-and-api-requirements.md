
# <a name="specify-office-hosts-and-api-requirements"></a>Spécification des exigences en matière d’hôtes Office et d’API



Il se peut que votre complément Office dépende d’un hôte Office spécifique, d’un ensemble de conditions requises, d’un membre d’API ou d’une version de l’API pour fonctionner correctement. Par exemple, votre complément peut :

- exécuter une ou plusieurs application Office (Word ou Excel) ;
    
- utiliser des API JavaScript disponibles uniquement dans certaines versions d’Office. Par exemple, vous pouvez utiliser les API JavaScript d’Excel dans un complément qui fonctionne dans Excel 2016 ; 
    
- s’exécuter uniquement dans les versions d’Office qui prennent en charge les membres d’API utilisés par votre complément.
    
Cet article vous aidera à comprendre les options que vous devez choisir afin de vous assurer que votre complément fonctionne comme prévu et atteint l’audience la plus large possible.

>**Remarque :** Pour savoir de manière détaillée quelle version d’Office prend en charge les compléments Office, consultez la page relative à la [disponibilité des compléments Office sur les plateformes et les hôtes](http://dev.office.com/add-in-availability). 

Le tableau suivant répertorie les concepts principaux décrits dans cet article.


|**Concept**|**Description**|
|:-----|:-----|
|Application Office, application hôte Office ou hôte Office|Application Office utilisée pour exécuter votre complément. Par exemple, Word, Word Online ou Excel.|
|Plateforme|Application sur laquelle l’hôte Office est exécuté, comme Office Online ou Office pour iPad.|
|Ensemble de conditions requises|Groupe nommé de membres d’API associés. Les compléments utilisent des ensembles de conditions requises pour déterminer si l’hôte Office prend en charge les membres d’API utilisés par votre complément. Il est plus facile de tester la prise en charge d’un ensemble de conditions requises, plutôt que la prise en charge de membres individuels d’API. La prise en charge de l’ensemble des conditions requises varie selon l’hôte Office et la version de ce dernier. <br >Les ensembles de conditions requises sont spécifiés dans le fichier manifeste. Quand vous définissez des ensembles de conditions requises dans le fichier manifeste, vous définissez le niveau minimal de prise en charge de l’API que l’hôte Office doit fournir pour exécuter votre complément. Les hôtes Office qui ne prennent pas en charge les ensembles de conditions requises spécifiés dans le manifeste ne peuvent pas exécuter votre complément, et votre complément ne sera pas affiché dans <span class="ui">Mes compléments</span>. Cela limite les emplacements où votre complément sera disponible. Dans le code utilisant les vérifications à l’exécution. Pour obtenir la liste complète des ensembles de conditions requises, voir [Ensemble de conditions requises pour les compléments Office](../../reference/office-add-in-requirement-sets.md).|
|Vérification à l’exécution|Test effectué à l’exécution pour déterminer si l’hôte Office qui exécute votre complément prend en charge les ensembles de conditions requises ou les méthodes utilisés par votre complément. Pour effectuer une vérification à l’exécution, vous pouvez utiliser une instruction **if** avec la méthode **isSetSupported**, les ensembles de conditions requises ou les noms de méthode qui ne font pas partie d’un ensemble de conditions requises. Les vérifications à l’exécution permettent de veiller à ce que votre complément atteigne le plus grand nombre possible de clients. Contrairement aux ensembles de conditions requises, les vérifications à l’exécution ne précisent pas le niveau minimal de prise en charge de l’API que l’hôte Office doit fournir pour l’exécution de votre complément. Au lieu de cela, vous devez utiliser l’instruction **if** afin de déterminer si un membre d’API est pris en charge. Si c’est le cas, vous pouvez fournir des fonctionnalités supplémentaires dans votre complément. Votre complément s’affiche toujours dans **Mes compléments** quand vous effectuez des vérifications à l’exécution.|

## <a name="before-you-begin"></a>Avant de commencer

Votre complément doit utiliser la version la plus récente du schéma de manifeste de complément. Si vous utilisez les vérifications à l’exécution dans votre complément, assurez-vous que vous utilisez la dernière API JavaScript pour la bibliothèque Office (office.js).


### <a name="specify-the-latest-add-in-manifest-schema"></a>Indication du schéma de manifeste de complément le plus récent

Le manifeste de votre du complément doit utiliser la version 1.1 du schéma de manifeste de complément. Définissez l’élément **App_office** dans votre manifeste complément comme suit.


```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```


### <a name="specify-the-latest-javascript-api-for-office-library"></a>Indication de l’API JavaScript la plus récente pour la bibliothèque Office


Si vous utilisez des vérifications à l’exécution, référencez la version la plus récente de l’API JavaScript pour la bibliothèque Office à partir du réseau de livraison de contenu (CDN). Pour ce faire, ajoutez la balise `script` suivante à votre code HTML. L’utilisation de `/1/` dans l’URL CDN garantit que vous référencez la version d’Office.js la plus récente.


```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```


## <a name="options-to-specify-office-hosts-or-api-requirements"></a>Options pour spécifier des hôtes Office ou les conditions requises d’API

Lors de la spécification des hôtes Office ou des conditions requises d’API, vous devez tenir compte de plusieurs facteurs. Le diagramme suivant montre comment choisir la technique à utiliser dans votre complément.


![Optez pour la meilleure solution pour votre complément lorsque vous spécifiez des hôtes Office ou des exigences d’API](../../images/e3498f8f-7c7c-461c-84f3-b93910b088b9.png)

- Si votre complément s’exécute dans un hôte Office, définissez l’élément **Hosts** dans le manifeste. Pour plus d’informations, voir [Définition de l’élément Hosts](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-hosts-element).
    
- Pour définir l’ensemble minimal de conditions requises ou les membres minimaux d’API qu’un hôte Office doit prendre en charge pour exécuter votre complément, définissez l’élément **Requirements** dans le manifeste. Pour plus d’informations, consultez la section [ Définition de l’élément Requirements dans le manifeste](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).
    
- Si vous souhaitez proposer des fonctionnalités supplémentaires lorsque des ensembles de conditions requises spécifiques ou des membres d’API sont disponibles dans l’hôte Office, effectuez une vérification à l’exécution dans le code JavaScript de votre complément. Par exemple, si votre complément est exécuté dans Excel 2016, utilisez les membres d’API de la nouvelle API JavaScript pour Excel pour fournir des fonctionnalités supplémentaires. Pour plus d’informations, consultez la section [Utilisation des vérifications à l’exécution dans votre code JavaScript](../../docs/overview/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).
    

## <a name="set-the-hosts-element"></a>Définition de l’élément Hosts


Pour exécuter votre complément dans une application hôte Office, utilisez les éléments **Hosts** et **Host** dans le manifeste. Si vous ne définissez pas l’élément **Hosts**, votre complément sera exécuté dans tous les hôtes.

Par exemple, les déclarations  **Hosts** et **Host** suivantes indiquent que le complément fonctionnera avec n’importe quelle version d’Excel, y compris Excel pour Windows, Excel Online et Excel pour iPad.




```XML
  <Hosts>
    <Host Name="Workbook" />
  </Hosts>
```

L’élément  **Hosts** peut contenir un ou plusieurs éléments  **Host**. L’élément  **Host** indique l’hôte Office requis par votre complément. L’attribut **Name** est requis et peut être défini sur l’une des valeurs suivantes.



| Name          | Applications hôtes Office                      |
|:--------------|:----------------------------------------------|
| Base de données      | applications web Access                               |
| Document      | Word pour Windows, Mac, iPad et Online        |
| Boîte aux lettres       | Outlook pour Windows, Mac, Web et Outlook.com | 
| Présentation  | PowerPoint pour Windows, Mac, iPad et Online  |
| Projet       | Projet                                       |
| Classeur      | Excel pour Windows, Mac, iPad et Online           |

 >**Remarque :**  L’attribut `Name` spécifie l’application hôte Office pouvant exécuter votre complément. Les hôtes Office sont pris en charge sur différentes plateformes et sont exécutés sur les ordinateurs de bureau, les navigateurs web, les tablettes et les appareils mobiles. Vous ne pouvez pas indiquer quelle plateforme peut être utilisée pour exécuter votre complément. Par exemple, si vous spécifiez `Mailbox`, Outlook et Outlook Web App peuvent être utilisés pour exécuter votre complément. 


## <a name="set-the-requirements-element-in-the-manifest"></a>Définition de l’élément Requirements dans le manifeste


L’élément **Requirements** indique les ensembles de conditions minimales requises ou les membres d’API qui doivent être pris en charge par l’hôte Office en vue d’exécuter votre complément. L’élément **Requirements** peut indiquer des ensembles de conditions requises et des méthodes individuelles utilisés dans votre complément. Dans la version 1.1 du schéma de manifeste du complément, l’élément **Requirements** est facultatif pour tous les compléments, sauf pour les compléments Outlook.


 >**Attention :**  utilisez uniquement l’élément **Requirements** pour spécifier des ensembles de conditions requises essentiels ou des membres API que votre complément doit utiliser. Si la plateforme ou l’hôte Office ne prend pas en charge les ensembles de conditions requises ou les membres d’API spécifiés dans l’élément **Requirements**, le complément ne s’exécute pas dans cet hôte ou cette plateforme et ne s’affiche pas dans **Mes compléments**. Nous vous recommandons plutôt de rendre votre complément disponible sur toutes les plateformes d’un hôte Office, comme Excel pour Windows, Excel Online et Excel pour iPad. Pour rendre votre complément disponible sur _tous_ les hôtes et plateformes Office, utilisez des vérifications à l’exécution à la place de l’élément **Requirements**.

Cet exemple de code illustre un complément qui se charge dans toutes les applications hôtes Office qui prennent en charge les éléments suivants :


-  Un ensemble de conditions requises **TableBindings**, dont la version minimale est 1.1.
    
-  Un ensemble de conditions requises **OOXML**, dont la version minimale est 1.1.
    
-  La méthode **Document.getSelectedDataAsync**.
    



```XML
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" MinVersion="1.1"/>
      <Set Name="OOXML" MinVersion="1.1"/>
   </Sets>
   <Methods>
      <Method Name="Document.getSelectedDataAsync"/>
   </Methods>
</Requirements>
```

- L’élément **Requirements** contient les éléments enfants **Sets** et **Methods**.
    
- L’élément  **Sets** peut contenir un ou plusieurs éléments  **Set**.  **DefaultMinVersion** indique la valeur **MinVersion** par défaut de tous les éléments  **Set** enfants.
    
- L’élément **Set** spécifie les ensembles de conditions requises que l’hôte Office doit prendre en charge pour exécuter le complément. L’attribut **Name** indique le nom de l’ensemble de conditions requises. L’attribut **MinVersion** spécifie la version minimale de l’ensemble de conditions requises. L’attribut **MinVersion** remplace la valeur de **DefaultMinVersion**. Pour plus d’informations sur les ensembles de conditions requises et les versions auxquelles les membres de votre API appartiennent, consultez [Ensemble de conditions requises pour les compléments Office](../../reference/office-add-in-requirement-sets.md).
    
- L’élément **Methods** peut contenir un ou plusieurs éléments **Method**. Vous ne pouvez pas utiliser l’élément **Methods** avec des compléments Outlook.
    
- L’élément  **Method** spécifie une méthode individuelle qui doit être prise en charge dans l’hôte Office où votre complément est exécuté. L’attribut **Name** est obligatoire et indique le nom de la méthode qualifiée avec son objet parent.
    

## <a name="use-runtime-checks-in-your-javascript-code"></a>Utilisation des vérifications à l’exécution dans votre code JavaScript


Vous pouvez fournir des fonctionnalités supplémentaires dans votre complément si certains ensembles de conditions requises sont pris en charge par l’hôte Office. Par exemple, vous pouvez utiliser les nouvelles interfaces API JavaScript de Word dans votre complément existant si ce dernier est exécuté dans Word 2016. Pour ce faire, utilisez la méthode **isSetSupported** avec le nom de l’ensemble de conditions requises. **isSetSupported** détermine, lors de l’exécution, si l’hôte Office exécutant le complément prend en charge l’ensemble des conditions requises. Si l’ensemble de conditions requises est pris en charge, **isSetSupported** renvoie **True** et exécute le code supplémentaire qui utilise les membres d’API provenant de l’ensemble de conditions requises. Si l’hôte Office ne prend pas en charge l’ensemble de conditions requises, **isSetSupported** renvoie **False** et le code supplémentaire n’est pas exécuté. Le code suivant indique la syntaxe à utiliser avec **isSetSupported**.


```js
if (Office.context.requirements.isSetSupported(RequirementSetName , VersionNumber )
{
   // Code that uses API members from RequirementSetName .
}

```


-  _RequirementSetName_ (obligatoire) est une chaîne représentant le nom de l’ensemble de conditions requises. Pour plus d’informations sur les ensembles de conditions requises disponibles, voir [Ensemble de conditions requises pour les compléments Office](../../reference/office-add-in-requirement-sets.md).
    
-  _VersionNumber_ (facultatif) correspond à la version de l’ensemble de conditions requises.
    
Dans Excel 2016 ou Word 2016, utilisez **isSetSupported** avec les ensembles de conditions requises  **ExcelAPI** ou **WordAPI**. La méthode  **isSetSupported**, ainsi que les ensembles de conditions requises  **ExcelAPI** et **WordAPI**, sont disponibles dans le dernier fichier Office.js du CDN. Si vous n’utilisez pas Office.js à partir du CDN, votre complément peut générer des exceptions, car la méthode  **isSetSupported** ne sera pas définie. Pour plus d’informations, voir [ Indication de l’API JavaScript la plus récente pour la bibliothèque Office](../../docs/overview/specify-office-hosts-and-api-requirements.md#specify-the-latest-javascript-api-for-office-library). 


 >**Remarque :**   **isSetSupported** ne fonctionne pas dans Outlook ou Outlook Web App. Pour utiliser une vérification à l’exécution dans Outlook ou Outlook Web App, utilisez la technique décrite dans la section [Vérifications à l’exécution à l’aide de méthodes ne faisant pas partie d’un ensemble de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md#runtime-checks-using-methods-not-in-a-requirement-set).

L’exemple de code suivant montre comment un complément peut fournir des fonctionnalités différentes pour divers hôtes Office qui peuvent prendre en charge plusieurs ensembles de conditions requises ou membres d’API.




```js
if (Office.context.requirements.isSetSupported('WordApi', 1.1)
{
       // Run code that provides additional functionality using the JavaScript API for Word when the add-in runs in Word 2016.
}
else if (Office.context.requirements.isSetSupported('CustomXmlParts')
{
      // Run code that uses API members from the CustomXmlParts requirement set.
}
else 
{
    // Run additional code when the Office host is not Word 2016, and when the Office host does not support the CustomXmlParts requirement set.
}

```


## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a>Vérifications à l’exécution à l’aide de méthodes ne faisant pas partie d’un ensemble de conditions requises


Certains membres API n’appartiennent pas à des ensembles de conditions requises. Cela s’applique uniquement aux membres d’API qui font partie de l’espace de noms de l’[interface API JavaScript pour Office](../../reference/javascript-api-for-office.md) (rien sous Office), et non aux membres d’API qui appartiennent à l’espace de noms de l’interface API JavaScript pour Word (rien dans Word) ou de la [référence de l’API JavaScript pour les compléments Excel](https://msdn.microsoft.com/library/office/mt616490.aspx) (rien dans Excel). Lorsque votre complément dépend d’une méthode qui ne fait pas partie d’un ensemble de conditions requises, vous pouvez utiliser la vérification à l’exécution pour déterminer si la méthode est prise en charge par l’hôte Office, comme indiqué dans l’exemple suivant. Pour consulter la liste complète des méthodes qui n’appartiennent pas à un ensemble de conditions requises, voir [Ensemble de conditions requises pour les compléments Office](../../reference/office-add-in-requirement-sets.md).


 >**Remarque** : nous vous recommandons de limiter l’utilisation de ce type de vérification à l’exécution dans le code de votre complément.

L’exemple de code suivant vérifie si l’hôte prend en charge **document.setSelectedDataAsync**.




```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses document.setSelectedDataAsync.
}
```


## <a name="additional-resources"></a>Ressources supplémentaires



- [Manifeste XML des compléments Office](../../docs/overview/add-in-manifests.md)
    
- [Ensembles de conditions requises pour les compléments Office](../../reference/requirement-sets/office-add-in-requirement-sets.md)
    
- [Word-Add-in-Get-Set-EditOpen-XML ](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
    
