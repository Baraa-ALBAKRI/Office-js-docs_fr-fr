# <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Spécifier les hôtes Office et les conditions requises d’API](../docs/overview/specify-office-hosts-and-api-requirements.md).

Pour plus d’informations sur la prise en charge des compléments par l’hôte Office, reportez-vous à la rubrique [Disponibilité des compléments Office sur les plateformes et les hôtes](https://dev.office.com/add-in-availability).

## <a name="hostspecific-api-requirement-sets"></a>Ensembles de conditions requises de l’API spécifique à l’hôte

Pour plus d’informations sur les ensembles de conditions requises des API pour Excel, Word, OneNote, Outlook et Dialog, reportez-vous à :

- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md)
- [Ensembles de conditions requises de l’API JavaScript pour OneNote](onenote-api-requirement-sets.md)
- [Présentation de l’ensemble de conditions requises pour les API Outlook](../outlook/tutorial-api-requirement-sets.md)
[Ensembles de conditions requises de l’API de boîte de dialogue](dialog-api-requirement-sets.md)

## <a name="common-api-requirement-sets"></a>Ensembles de conditions requises des API communes

Le tableau suivant répertorie les ensembles de conditions requises communs, les méthodes de chaque ensemble et les applications hôtes Office qui les prennent en charge. Tous ces ensembles de conditions requises d’API sont à la version 1.1.


|  Ensemble de conditions requises  |  Hôte Office  |  Méthodes dans l’ensemble  |
|:-----|-----|:-----|:-----|
| ActiveView | PowerPoint<br>PowerPoint&nbsp;Online|Document.getActiveViewAsync|
| BindingEvents  | Applications web Access<br>Excel<br>Excel Online<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | PowerPoint<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad<br/>Excel Online<br/>PowerPoint Online|Prend en charge la sortie au format Office Open XML (OOXML) sous la forme d’un tableau d’octets<br>(Office.FileType.Compressed) lorsque vous utilisez la méthode Document.getFileAsync.|
| CustomXmlParts    | Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DocumentEvents    | Excel<br>Excel Online<br>PowerPoint Online<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| Fichier  | PowerPoint<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad<br>PowerPoint Online|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Prend en charge le forçage au format HTML (Office.CoercionType.Html) lors de la lecture et de l’écriture des données à l’aide des méthodes Document.getSelectedDataAsync,<br>Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| ImageCoercion | Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Prise en charge de la conversion en une image (Office.CoercionType.Image) lors de l’écriture des données à l’aide de la méthode Document.setSelectedDataAsync.|
| Boîte aux lettres   |Outlook pour Windows<br>Outlook pour le web<br>Outlook pour Mac<br>Outlook Web App |Voir [Présentation de l’ensemble de conditions requises pour les API Outlook](./outlook/tutorial-api-requirement-sets.md).|
| MatrixBindings    | Excel<br>Excel Online<br>Word<br>Word Online|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | Excel<br>Excel Online<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Prise en charge du forçage de type sur la structure de données (Office.CoercionType.Matrix) « matrice » (tableau de tableaux) lors de la lecture et de l’écriture de données à l’aide des méthodes Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| OoxmlCoercion | Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Prise en charge du forçage de type au format Open Office XML (OOXML) (Office.CoercionType.Ooxml) lors de la lecture et de l’écriture de données à l’aide des méthodes Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| PartialTableBindings  | Applications web Access||
| PdfFile   | PowerPoint<br/>PowerPoint Online<br/>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Prend en charge la sortie au format PDF (Office.FileType.Pdf)<br>lorsque vous utilisez la méthode Document.getFileAsync.|
| Sélection | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Paramètres  | Applications web Access<br>Excel<br>Excel Online<br>PowerPoint<br>PowerPoint Online<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | Applications web Access<br>Excel<br>Excel Online<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | Applications web Access<br>Excel<br>Excel Online<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Prise en charge du forçage de type sur la structure de données « tableau » (Office.CoercionType.Table) lors de la lecture et de l’écriture de données à l’aide des méthodes Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| TextBindings  | Excel<br>Excel Online<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Prise en charge du forçage de type au format texte (Office.CoercionType.Text) lors de la lecture et de l’écriture de données à l’aide des méthodes Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| TextFile  | Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad<br/>|Prise en charge de sortie au format texte (Office.FileType.Text) lors de l’utilisation de la méthode Document.getFileAsync.|

## <a name="methods-that-arent-part-of-a-requirement-set"></a>Méthodes qui ne font pas partie d’un ensemble de conditions requises

Les méthodes suivantes dans l’interface API JavaScript pour Office ne font pas partie d’un ensemble de conditions requises. Si l’une de ces méthodes est nécessaire pour votre complément, utilisez les éléments **Methods** et **Method** dans le manifeste du complément afin de déclarer qu’elles sont obligatoires ou effectuez la vérification à l’exécution à l’aide d’une instruction if. Pour plus d’informations, voir l’article sur la [spécification des conditions requises pour les API et les hôtes Office](../docs/overview/specify-office-hosts-and-api-requirements.md).

|**Nom de la méthode**|**Prise en charge des hôtes Office**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Applications web Access, Excel et Excel Online|
|Document.getFilePropertiesAsync|Excel, Excel Online, Word et PowerPoint|
|Document.getProjectFieldAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getResourceFieldAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getSelectedViewAsync|PowerPoint et PowerPoint Online|
|Document.getTaskAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getTaskFieldAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.goToByIdAsync|Excel, Excel Online, Word et PowerPoint|
|Settings.addHandlerAsync|Applications web Access, Excel, Excel Online, Word et PowerPoint|
|Settings.refreshAsync|Applications web Access, Excel, Excel Online, Word, PowerPoint et PowerPoint Online|
|Settings.removeHandlerAsync|Applications web Access, Excel, Excel Online, Word et PowerPoint|
|TableBinding.clearFormatsAsync|Excel, Excel Online|
|TableBinding.setFormatsAsync|Excel, Excel Online|
|TableBinding.setTableOptionsAsync|Excel, Excel Online|

## <a name="additional-resources"></a>Ressources supplémentaires

- [Spécification des exigences en matière d’hôtes Office et d’API](../docs/overview/specify-office-hosts-and-api-requirements.md)



