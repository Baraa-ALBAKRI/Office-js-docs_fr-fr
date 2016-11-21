# <a name="excel-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour Excel

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Spécifier les hôtes Office et les conditions requises d’API](../docs/overview/specify-office-hosts-and-api-requirements.md).

Les compléments Excel peuvent être exécutés dans différentes versions d’Office, notamment Office 2016 pour Windows, Office pour iPad, Office pour Mac et Office Online. Le tableau suivant répertorie les ensembles de conditions requises pour Excel, les applications hôtes Office qui prennent en charge ces conditions et la version ou le numéro de build de ces applications. 

|  Ensemble de conditions requises  |  Office 2016 pour Windows*  |  Office 2016 pour iPad  |  Office 2016 pour Mac  | Office Online  |
|:-----|-----|:-----|:-----|:-----|
| ExcelApi 1.3  | Version 1608 ou versions ultérieures| 1.27 ou version ultérieure |  15.27 ou version ultérieure| Septembre 2016 | 
| ExcelApi 1.2  | Version 1601 ou versions ultérieures | 1.21 ou version ultérieure | 15.22 ou version ultérieure| Janvier 2016 |
| ExcelApi 1.1  | Version 1509 (Build 4266.1001) ou version ultérieure | 1.19 ou version ultérieure | 15.20 ou version ultérieure| Janvier 2016 |

> &#42; **Remarque** : Le numéro de build d’Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que l’ensemble de conditions requises de l’ExcelApi 1.1.

Pour en savoir plus sur les numéros de version et de build, voir :

- [Numéros de version et de build des canaux de réception des mises à jour pour les clients Office 365](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [Quelle est la version d’Office que j’utilise ?](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- [Où trouver le numéro de version et de build pour une application cliente Office 365](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office
Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="whats-new-in-excel-javascript-api-13"></a>Nouveautés de l’API JavaScript 1.3 pour Excel 
Les ajouts apportés aux API JavaScript pour Excel dans l’ensemble de conditions requises 1.3 sont présentés ci-dessous. 

|Objet| Nouveautés| Description|Ensemble de conditions requises|
|:----|:----|:----|:----|
|[binding](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/binding.md)|_Méthode_ > [delete()](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/binding.md#delete)|Supprime la liaison.|1.3|
|[bindingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md)|_Méthode_ > [add(range: Plage ou chaîne, bindingType: chaîne, id: chaîne)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md#addrange-range-or-string-bindingtype-string-id-string)|Ajouter une nouvelle liaison à une plage spécifique.|1.3|
|[bindingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md)|_Méthode_ > [addFromNamedItem (name: chaîne, bindingType: chaîne, id: chaîne)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md#addfromnameditemname-string-bindingtype-string-id-string)|Ajouter une nouvelle liaison basée sur un élément nommé dans le classeur.|1.3|
|[bindingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md)|_Méthode_ > [addFromSelection (bindingType: chaîne, id: chaîne)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md#addfromselectionbindingtype-string-id-string)|Ajouter un nouvelle liaison basée sur la sélection en cours.|1.3|
|[bindingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md)|_Méthode_ > [getItemOrNull(id: chaîne)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/bindingcollection.md#getitemornullid-string)|Obtient un objet de liaison par ID. Si l’objet de liaison n’existe pas, la propriété isNull de l’objet renvoyé aura la valeur true.|1.3|
|[chartCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/chartcollection.md)|_Méthode_ > [getItemOrNull(name: chaîne)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/chartcollection.md#getitemornullname-string)|Extrait un graphique à l’aide de son nom. Si plusieurs graphiques portent le même nom, c’est le premier d’entre eux qui est renvoyé.|1.3|
|[namedItemCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/nameditemcollection.md)|_Méthode_ > [getItemOrNull(name: chaîne)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/nameditemcollection.md#getitemornullname-string)|Obtient un objet NamedItem à l’aide de son nom. Si l’objet NamedItem n’existe pas, la propriété isNull de l’objet renvoyé aura la valeur true.|1.3|
|[pivotTable](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottable.md)|_Propriété_ > name|Nom du tableau croisé dynamique.|1.3|
|[pivotTable](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottable.md)|_Relation_ > worksheet|Feuille de calcul contenant le tableau croisé dynamique. En lecture seule.|1.3|
|[pivotTable](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottable.md)|_Méthode_ > [refresh()](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottable.md#refresh)|Actualise le tableau croisé dynamique.|1.3|
|[pivotTableCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottablecollection.md)|_Propriété_ > items|Collection d’objets de tableau croisé dynamique. En lecture seule.|1.3|
|[pivotTableCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottablecollection.md)|_Méthode_ > [getItem(name: chaîne)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottablecollection.md#getitemname-string)|Extrait un tableau croisé dynamique par nom.|1.3|
|[pivotTableCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottablecollection.md)|_Méthode_ > [getItemOrNull(name: chaîne)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/pivottablecollection.md#getitemornullname-string)|Extrait un tableau croisé dynamique par nom. Si le tableau croisé dynamique n’existe pas, la propriété isNull de l’objet renvoyé aura la valeur true.|1.3|
|[range](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/range.md)|_Méthode_ > [getIntersectionOrNull(anotherRange: Plage ou chaîne)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/range.md#getintersectionornullanotherrange-range-or-string)|Obtient l’objet de plage qui représente l’intersection rectangulaire des plages données. Si aucune intersection n’est trouvée, renvoie un objet Null.|1.3|
|[range](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/range.md)|_Méthode_ > [getVisibleView()](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/range.md#getvisibleview)|Représente les lignes visibles de la plage en cours.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propriété_ > cellAddresses|Représente les adresses de cellule de la RangeView. En lecture seule.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propriété_ > columnCount|Renvoie le nombre de colonnes visibles. En lecture seule.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propriété_ > formulas|Représente la formule dans le style de notation A1.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propriété_ > formulasLocal|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur.  Par exemple, la formule « =SUM(A1, présentée dans 1.5) » en anglais deviendrait « =SUMME(A1; 1,5) » en allemand.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propriété_ > formulasR1C1|Représente la formule dans le style de notation R1C1.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propriété_ > index|Renvoie une valeur qui représente l’index de l’affichage de plage. En lecture seule.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propriété_ > numberFormat|Représente le code de format de nombre d’Excel pour une cellule donnée.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propriété_ > rowCount|Renvoie le nombre de lignes visibles. En lecture seule.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propriété_ > text|Valeurs de texte de la plage spécifiée. La valeur de texte ne dépend pas de la largeur de la cellule. Le remplacement par le signe # qui se produit dans l’interface utilisateur d’Excel n’a aucun effet sur la valeur de texte renvoyée par l’API. En lecture seule.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propriété_ > valueTypes|Représente le type de données de chaque cellule. En lecture seule. Les valeurs possibles sont les suivantes : Unknown (inconnu), Empty (vide), String (chaîne), Integer (entier), Double (double), Boolean (valeur booléenne), Error (erreur).|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Propriété_ > values|Représente les valeurs brutes de l’affichage de plage spécifié. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Une cellule contenant une erreur renvoie la chaîne d’erreur.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Relation_ > rows|Représente une collection d’affichages de plage associés à la plage. En lecture seule.|1.3|
|[rangeView](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md)|_Méthode_ > [getRange()](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeview.md#getrange)|Obtient la plage parent associée à l’affichage de plage actuel.|1.3|
|[rangeViewCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeviewcollection.md)|_Propriété_ > items|Collection d’objets rangeView. En lecture seule.|1.3|
|[rangeViewCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeviewcollection.md)|_Methode_ > [getItemAt(index: nombre)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/rangeviewcollection.md#getitematindex-number)|Obtient une ligne d’affichage de plage par l’intermédiaire de son index. Avec index de base zéro.|1.3|
|[setting](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/setting.md)|_Propriété_ > key|Renvoie la clé qui représente l’id du paramètre. En lecture seule.|1.3|
|[setting](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/setting.md)|_Méthode_ > [delete()](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/setting.md#delete)|Supprime le paramètre.|1.3|
|[settingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md)|_Propriété_ > items|Collection d’objets setting. En lecture seule.|1.3|
|[settingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md)|_Méthode_ > [getItem(key: chaîne)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md#getitemkey-string)|Obtient une entrée Setting via la clé.|1.3|
|[settingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md)|_Méthode_ > [getItemOrNull(key: chaîne)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md#getitemornullkey-string)|Obtient une entrée Setting via la clé. Si l’objet Setting n’existe pas, la propriété isNull de l’objet renvoyé aura la valeur true.|1.3|
|[settingCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md)|_Méthode_ > [set(key: chaîne, value: chaîne)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingcollection.md#setkey-string-value-string)|Définit ou ajoute le paramètre spécifié dans le classeur.|1.3|
|[settingsChangedEventArgs](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/settingschangedeventargs.md)|_Relation_ > settingCollection|Obtient l’objet Setting qui représente la liaison qui a déclenché l’événement SettingsChanged.|1.3|
|[table](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/table.md)|_Propriété_ > highlightFirstColumn|Indique si la première colonne contient une mise en forme spéciale.|1.3|
|[table](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/table.md)|_Propriété_ > highlightLastColumn|Indique si la dernière colonne contient une mise en forme spéciale.|1.3|
|[table](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/table.md)|_Propriété_ > showBandedColumns|Indique si les colonnes affichent une mise en forme à bandes dans laquelle la mise en évidence des colonnes impaires diffère de celle des colonnes paires pour faciliter la lecture du tableau.|1.3|
|[table](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/table.md)|_Propriété_ > showBandedRows|Indique si les lignes affichent une mise en forme à bandes dans laquelle la mise en évidence des lignes impaires diffère de celle des lignes paires pour faciliter la lecture du tableau.|1.3|
|[table](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/table.md)|_Propriété_ > showFilterButton|Indique si les boutons de filtre sont visibles dans la partie supérieure de chaque en-tête de colonne. Ce paramètre est autorisé uniquement si le tableau contient une ligne d’en-tête.|1.3|
|[tableCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/tablecollection.md)|_Méthode_ > [getItemOrNull(key : nombre ou chaîne)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/tablecollection.md#getitemornullkey-number-or-string)|Obtient un tableau à l’aide de son nom ou de son ID. Si le tableau n’existe pas, la propriété isNull de l’objet renvoyé aura la valeur true.|1.3|
|[tableColumnCollection](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/tablecolumncollection.md)|_Méthode_ > [getItemOrNull(key : nombre ou chaîne)](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/tablecolumncollection.md#getitemornullkey-number-or-string)|Obtient un objet de colonne par son nom ou son ID. Si la colonne n’existe pas, la propriété isNull de l’objet renvoyé aura la valeur true.|1.3|
|[workbook](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/workbook.md)|_Relation_ > pivotTables|Représente une collection de tableaux croisés dynamiques associés au classeur. En lecture seule.|1.3|
|[workbook](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/workbook.md)|_Relation_ > settings|Représente une collection d’objets Settings associés au classeur. En lecture seule.|1.3|
|[worksheet](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec/reference/excel/worksheet.md)|_Relation_ > pivotTables|Collection de tableaux croisés dynamiques qui font partie de la feuille de calcul. En lecture seule.|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Nouveautés de l’API JavaScript 1.2 pour Excel
Les ajouts apportés aux API JavaScript pour Excel dans l’ensemble de conditions requises 1.2 sont présentés ci-dessous. 

|Objet| Nouveautés| Description|Ensemble de conditions requises|
|:----|:----|:----|:----|
|[chart](../excel/chart.md)|_Propriété_ > id|Extrait un graphique en fonction de sa position dans la collection. En lecture seule.|1.2|
|[chart](../excel/chart.md)|_Relation_ > worksheet|Feuille de calcul contenant le graphique actuel. En lecture seule.|1.2|
|[chart](../excel/chart.md)|_Méthode_ > [getImage(height: nombre, width: nombre, fittingMode: chaîne)](../excel/chart.md#getimageheight-number-width-number-fittingmode-string)|Affiche le graphique sous forme d’image codée en Base64 ajustée aux dimensions spécifiées.|1.2|
|[filter](../excel/filter.md)|_Relation_ > criteria|Le filtre actuellement appliqué à la colonne donnée. En lecture seule.|1.2|
|[filter](../excel/filter.md)|_Méthode_ > [apply(criteria: FilterCriteria)](../excel/filter.md#applycriteria-filtercriteria)|Appliquer les critères de filtre donnés à la colonne indiquée.|1.2|
|[filter](../excel/filter.md)|_Méthode_ > [applyBottomItemsFilter(count: nombre)](../excel/filter.md#applybottomitemsfiltercount-number)|Appliquer un filtre « Élément inférieur » à la colonne pour le nombre d’éléments donné.|1.2|
|[filter](../excel/filter.md)|_Méthode_ > [applyBottomPercentFilter(percent: nombre)](../excel/filter.md#applybottompercentfilterpercent-number)|Appliquer un filtre « Pourcentage inférieur » à la colonne pour le pourcentage d’éléments donné.|1.2|
|[filter](../excel/filter.md)|_Méthode_ > [applyCellColorFilter(color: chaîne)](../excel/filter.md#applycellcolorfiltercolor-string)|Appliquer un filtre « Couleur de cellule » à la colonne pour la couleur donnée.|1.2|
|[filter](../excel/filter.md)|_Méthode_ > [applyCustomFilter(criteria1: chaîne, criteria2: chaîne, oper: chaîne)](../excel/filter.md#applycustomfiltercriteria1-string-criteria2-string-oper-string)|Appliquer un filtre « Icône » à la colonne pour les chaînes de critères données.|1.2|
|[filter](../excel/filter.md)|_Méthode_ > [applyDynamicFilter(criteria: chaîne)](../excel/filter.md#applydynamicfiltercriteria-string)|Appliquer un filtre « Dynamique » à la colonne.|1.2|
|[filter](../excel/filter.md)|_Méthode_ > [applyFontColorFilter(color: chaîne)](../excel/filter.md#applyfontcolorfiltercolor-string)|Appliquer un filtre « Couleur de police » à la colonne pour la couleur donnée.|1.2|
|[filter](../excel/filter.md)|_Méthode_ > [applyIconFilter(icon: Icône)](../excel/filter.md#applyiconfiltericon-icon)|Appliquer un filtre « Icône » à la colonne pour l’icône donnée.|1.2|
|[filter](../excel/filter.md)|_Méthode_ > [applyTopItemsFilter(count: nombre)](../excel/filter.md#applytopitemsfiltercount-number)|Appliquer un filtre « Élément supérieur » à la colonne pour le nombre d’éléments donné.|1.2|
|[filter](../excel/filter.md)|_Méthode_ > [applyTopPercentFilter(percent: nombre)](../excel/filter.md#applytoppercentfilterpercent-number)|Appliquer un filtre « Pourcentage supérieur » à la colonne pour le pourcentage d’éléments donné.|1.2|
|[filter](../excel/filter.md)|_Méthode_ > [applyValuesFilter(values: ()[])](../excel/filter.md#applyvaluesfiltervalues-)|Appliquer un filtre « Valeurs » à la colonne pour les valeurs données.|1.2|
|[filter](../excel/filter.md)|_Méthode_ > [clear()](../excel/filter.md#clear)|Effacer le filtre sur la colonne donnée.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Propriété_ > color|Chaîne de couleur HTML utilisée pour filtrer des cellules. Utilisée avec le filtrage « cellColor » et « fontColor ».|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Propriété_ > criterion1|Premier critère utilisé pour filtrer des données. Utilisé comme opérateur dans le cas d’un filtrage « Custom ».|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Propriété_ > criterion2|Second critère utilisé pour filtrer des données. Utilisé uniquement comme opérateur dans le cas d’un filtrage « Custom ».|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Propriété_ > dynamicCriteria|Critères dynamiques de l’ensemble Excel.DynamicFilterCriteria à appliquer à cette colonne. Utilisé avec un filtrage « Dynamic ». Les valeurs possibles sont les suivantes : Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Propriété_ > filterOn|Propriété utilisée par le filtre pour déterminer si les valeurs doivent rester visibles. Les valeurs possibles sont les suivantes : BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Propriété_ > operator|Opérateur utilisé pour combiner les critères 1 et 2 lorsque vous utilisez le filtrage « Custom ». Les valeurs possibles sont les suivantes : And, Or.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Propriété_ > values|Valeurs à utiliser pour le filtrage « Values ».|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Relation_ > icon|Icône utilisée pour filtrer des cellules. Utilisé avec le filtrage « icon ».|1.2|
|[filterDatetime](../excel/filterdatetime.md)|_Propriété_ > date|Date au format ISO8601 utilisée pour filtrer des données.|1.2|
|[filterDatetime](../excel/filterdatetime.md)|_Propriété_ > specificity|Utilisation de la date pour conserver des données. Par exemple, si la date est 2005-04-02 et la spécificité est définie sur « mois », le filtre conservera toutes les lignes dont la date correspond au mois d’avril 2009. Les valeurs possibles sont les suivantes : Year (année), Monday (lundi), Day (jour), Hour (heure), Minute (minute), Second (seconde).|1.2|
|[formatProtection](../excel/formatprotection.md)|_Propriété_ > formulaHidden|Indique si Excel masque la formule des cellules dans la plage. Une valeur null indique que les paramètres de formule masquée ne sont pas les mêmes sur l’ensemble de la plage.|1.2|
|[formatProtection](../excel/formatprotection.md)|_Propriété_ > locked|Indique si Excel verrouille les cellules dans l’objet. Une valeur null indique que les paramètres de verrouillage ne sont pas les mêmes sur l’ensemble de la plage.|1.2|
|[icon](../excel/icon.md)|_Propriété_ > index|Représente l’index de l’icône dans l’ensemble donné.|1.2|
|[icon](../excel/icon.md)|_Propriété_ > set|Représente l’ensemble dont fait partie l’icône. Les valeurs possibles sont les suivantes : Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.2|
|[range](../excel/range.md)|_Propriété_ > columnHidden|Indique si toutes les colonnes de la plage active sont masquées.|1.2|
|[range](../excel/range.md)|_Propriété_ > formulasR1C1|Représente la formule dans le style de notation R1C1.|1.2|
|[range](../excel/range.md)|_Propriété_ > hidden|Indique si toutes les cellules de la plage active sont masquées. En lecture seule.|1.2|
|[range](../excel/range.md)|_Propriété_ > rowHidden|Indique si toutes les lignes de la plage active sont masquées.|1.2|
|[range](../excel/range.md)|_Relation_ > sort|Représente le tri de plage de la plage actuelle. En lecture seule.|1.2|
|[range](../excel/range.md)|_Méthode_ > [merge(across: bool)](../excel/range.md#mergeacross-bool)|Fusionne la plage de cellules dans une zone de la feuille de calcul.|1.2|
|[range](../excel/range.md)|_Méthode_ > [unmerge()](../excel/range.md#unmerge)|Annule la fusion de la plage de cellules.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Propriété_ > columnWidth|Obtient ou définit la largeur de toutes les colonnes de la plage. Si les largeurs de colonne ne sont pas uniformes, la valeur « null » est renvoyée.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Propriété_ > rowHeight|Obtient ou définit la hauteur de toutes les lignes de la plage. Si les hauteurs de lignes ne sont pas uniformes, la valeur « null » est renvoyée.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Relation_ > protection|Renvoie l’objet de protection du format pour une plage. En lecture seule.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Méthode_ > [autofitColumns()](../excel/rangeformat.md#autofitcolumns)|Modifie la largeur des colonnes de la plage active pour obtenir le meilleur ajustement, en fonction des données présentes dans les colonnes.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Méthode_ > [autofitRows()](../excel/rangeformat.md#autofitrows)|Modifie la hauteur des lignes de la plage active pour obtenir le meilleur ajustement, en fonction des données présentes dans les colonnes.|1.2|
|[rangeReference](../excel/rangereference.md)|_Propriété_ > address|Représente les lignes visibles de la plage en cours.|1.2|
|[rangeSort](../excel/rangesort.md)|_Méthode_ > [apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: chaîne, method: chaîne)](../excel/rangesort.md#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|Effectue une opération de tri.|1.2|
|[sortField](../excel/sortfield.md)|_Propriété_ > ascending|Indique si le tri s’effectue dans l’ordre croissant.|1.2|
|[sortField](../excel/sortfield.md)|_Propriété_ > color|Couleur ciblée par la condition si le tri est appliqué à la couleur ou à la police de la cellule.|1.2|
|[sortField](../excel/sortfield.md)|_Propriété_ > dataOption|Options de tri supplémentaires pour ce champ. Les valeurs possibles sont les suivantes : Normal, TextAsNumber.|1.2|
|[sortField](../excel/sortfield.md)|_Propriété_ > key|Colonne (ou ligne, selon l’orientation du tri) ciblée par la condition. Représentée sous forme d’un décalage par rapport à la première colonne (ou ligne).|1.2|
|[sortField](../excel/sortfield.md)|_Propriété_ > sortOn|Type de tri de cette condition. Les valeurs possibles sont les suivantes : Value, CellColor, FontColor, Icon.|1.2|
|[sortField](../excel/sortfield.md)|_Relation_ > icon|Représente l’icône ciblée par la condition si le tri est appliqué à l’icône de la cellule.|1.2|
|[table](../excel/table.md)|_Relation_ > sort|Représente le tri du tableau. En lecture seule.|1.2|
|[table](../excel/table.md)|_Relation_ > worksheet|Feuille de calcul contenant le tableau actuel. En lecture seule.|1.2|
|[table](../excel/table.md)|_Méthode_ > [clearFilters()](../excel/table.md#clearfilters)|Supprime tous les filtres appliqués actuellement sur le tableau.|1.2|
|[table](../excel/table.md)|_Méthode_ > [convertToRange()](../excel/table.md#converttorange)|Convertit le tableau en plage normale de cellules. Toutes les données sont conservées.|1.2|
|[table](../excel/table.md)|_Méthode_ > [reapplyFilters()](../excel/table.md#reapplyfilters)|Applique de nouveau tous les filtres actuellement appliqués sur le tableau.|1.2|
|[tableColumn](../excel/tablecolumn.md)|_Relation_ > filter|Extrait le filtre appliqué à la colonne. En lecture seule.|1.2|
|[tableSort](../excel/tablesort.md)|_Propriété_ > matchCase|Indique si la casse a influé sur le dernier tri du tableau. En lecture seule.|1.2|
|[tableSort](../excel/tablesort.md)|_Propriété_ > method|Dernière méthode de classement des caractères chinois utilisée pour trier le tableau. En lecture seule. Les valeurs possibles sont les suivantes : PinYin, StrokeCount|1.2|
|[tableSort](../excel/tablesort.md)|_Relation_ > fields|Représente les dernières conditions utilisées pour trier le tableau. En lecture seule.|1.2|
|[tableSort](../excel/tablesort.md)|_Méthode_ > [apply(fields: SortField[], matchCase: bool, method: chaîne)](../excel/tablesort.md#applyfields-sortfield-matchcase-bool-method-string)|Effectue une opération de tri.|1.2|
|[tableSort](../excel/tablesort.md)|_Méthode_ > [clear()](../excel/tablesort.md#clear)|Efface le tri actuellement appliqué au tableau. Même si le classement du tableau n’est pas modifié, l’état des boutons d’en-tête est rétabli.|1.2|
|[tableSort](../excel/tablesort.md)|_Méthode_ > [reapply()](../excel/tablesort.md#reapply)|Applique à nouveau les paramètres actuels de tri au tableau.|1.2|
|[workbook](../excel/workbook.md)|_Relation_ > functions|Représente l’instance de l’application Excel contenant ce classeur. En lecture seule.|1.2|
|[worksheet](../excel/worksheet.md)|_Relation_ > protection|Renvoie un objet de protection de feuille pour une feuille de calcul. En lecture seule.|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_Propriété_ > protected|Indique si la feuille de calcul est protégée. En lecture seule. En lecture seule.|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_Relation_ > options|Options de protection de feuille. En lecture seule.|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_Méthode_ > [protect(options: WorksheetProtectionOptions)](../excel/worksheetprotection.md#protectoptions-worksheetprotectionoptions)|Protège une feuille de calcul. Échoue si la feuille de calcul est protégée.|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_Méthode_ > [unprotect()](../excel/worksheetprotection.md#unprotect)|Annule la protection d’une feuille de calcul.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propriété_ > allowAutoFilter|Représente l’option de protection de feuille de calcul qui autorise l’utilisation de la fonctionnalité Filtre automatique.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propriété_ > allowDeleteColumns|Représente l’option de protection de feuille de calcul qui autorise la suppression des colonnes.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propriété_ > allowDeleteRows|Représente l’option de protection de feuille de calcul qui autorise la suppression des lignes.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propriété_ > allowFormatCells|Représente l’option de protection de feuille de calcul qui autorise la mise en forme des cellules.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propriété_ > allowFormatColumns|Représente l’option de protection de feuille de calcul qui autorise la mise en forme des colonnes.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propriété_ > allowFormatRows|Représente l’option de protection de feuille de calcul qui autorise la mise en forme des lignes.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propriété_ > allowInsertColumns|Représente l’option de protection de feuille de calcul qui autorise l’insertion des colonnes.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propriété_ > allowInsertHyperlinks|Représente l’option de protection de feuille de calcul qui autorise l’insertion des liens hypertexte.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propriété_ > allowInsertRows|Représente l’option de protection de feuille de calcul qui autorise l’insertion des lignes.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propriété_ > allowPivotTables|Représente l’option de protection de feuille de calcul qui autorise l’utilisation de la fonctionnalité Tableau croisé dynamique.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Propriété_ > allowSort|Représente l’option de protection de feuille de calcul qui autorise l’utilisation de la fonctionnalité Tri.|1.2|

## <a name="excel-javascript-api-11"></a>API JavaScript 1.1 pour Excel
L’API JavaScript 1.1 pour Excel est la première version de l’API. Pour plus d’informations sur l’API, consultez les rubriques de référence sur l’API JavaScript pour Excel.  
    
## <a name="additional-resources"></a>Ressources supplémentaires

- [Spécification des exigences en matière d’hôtes Office et d’API](../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](https://dev.office.com/docs/add-ins/overview/add-in-manifests)