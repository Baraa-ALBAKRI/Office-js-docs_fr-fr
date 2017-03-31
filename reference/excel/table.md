# <a name="table-object-javascript-api-for-excel"></a>Objet Table (API JavaScript pour Excel)

Représente un tableau Excel.

## <a name="properties"></a>Propriétés

| Propriété       | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|highlightFirstColumn|bool|Indique si la première colonne contient une mise en forme spéciale.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|highlightLastColumn|bool|Indique si la dernière colonne contient une mise en forme spéciale.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|id|int|Renvoie une valeur qui identifie le tableau dans un classeur donné. La valeur de l’identificateur reste identique, même lorsque le tableau est renommé. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Nom du tableau.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showBandedColumns|bool|Indique si les colonnes affichent une mise en forme à bandes dans laquelle la mise en évidence des colonnes impaires diffère de celle des colonnes paires pour faciliter la lecture du tableau.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showBandedRows|bool|Indique si les lignes affichent une mise en forme à bandes dans laquelle la mise en évidence des lignes impaires diffère de celle des lignes paires pour faciliter la lecture du tableau.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showFilterButton|bool|Indique si les boutons de filtre sont visibles dans la partie supérieure de chaque en-tête de colonne. Ce paramètre est autorisé uniquement si le tableau contient une ligne d’en-tête.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showHeaders|bool|Indique si la ligne d’en-tête est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne d’en-tête.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showTotals|bool|Indique si la ligne de total est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne de total.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|style|string|Valeur de constante qui représente le style du tableau. Les valeurs possibles sont les suivantes : TableStyleLight1 à TableStyleLight21, TableStyleMedium1 à TableStyleMedium28, TableStyleStyleDark1 à TableStyleStyleDark11. Vous pouvez également indiquer un style personnalisé présent dans le classeur.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_Consultez la remarque importante sur les performances liées au tableau avec les [formules](#setting-formulas)_



## <a name="relationships"></a>Relations
| Relation | Type    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|colonnes|[TableColumnCollection](tablecolumncollection.md)|Représente une collection de toutes les colonnes du tableau. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|Objet Rows|[TableRowCollection](tablerowcollection.md)|Représente une collection de toutes les lignes du tableau. En lecture seule.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|tri|[TableSort](tablesort.md)|Représente le tri du tableau. En lecture seule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|feuille de calcul|[Worksheet](worksheet.md)|Feuille de calcul contenant le tableau actif. En lecture seule.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description| Dem. Set|
|:---------------|:--------|:----------|:----|
|[clearFilters()](#clearfilters)|void|Supprime tous les filtres appliqués actuellement sur le tableau.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[convertToRange()](#converttorange)|[Range](range.md)|Convertit le tableau en plage normale de cellules. Toutes les données sont conservées.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[delete()](#delete)|void|Supprime le tableau.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Obtient l’objet de plage associé au corps de données du tableau.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Obtient l’objet de plage associé à la ligne d’en-tête du tableau.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|Renvoie l’objet de plage associé à l’intégralité du tableau.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Renvoie l’objet de plage associé à la ligne de total du tableau.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[reapplyFilters()](#reapplyfilters)|void|Applique de nouveau tous les filtres actuellement appliqués sur le tableau.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>Détails des méthodes


### <a name="clearfilters"></a>clearFilters()
Supprime tous les filtres appliqués actuellement sur le tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableObject.clearFilters();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

### <a name="converttorange"></a>convertToRange()
Convertit le tableau en plage normale de cellules. Toutes les données sont conservées.

#### <a name="syntax"></a>Syntaxe
```js
tableObject.convertToRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.convertToRange();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="delete"></a>delete()
Supprime le tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableObject.delete();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getdatabodyrange"></a>getDataBodyRange()
Obtient l’objet de plage associé au corps de données du tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableObject.getDataBodyRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableDataRange = table.getDataBodyRange();
    tableDataRange.load('address')
    return ctx.sync().then(function() {
            console.log(tableDataRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getheaderrowrange"></a>getHeaderRowRange()
Obtient l’objet de plage associé à la ligne d’en-tête du tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableObject.getHeaderRowRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableHeaderRange = table.getHeaderRowRange();
    tableHeaderRange.load('address');
    return ctx.sync().then(function() {
        console.log(tableHeaderRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrange"></a>getRange()
Renvoie l’objet de plage associé à l’intégralité du tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableObject.getRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableRange = table.getRange();
    tableRange.load('address');    
    return ctx.sync().then(function() {
            console.log(tableRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="gettotalrowrange"></a>getTotalRowRange()
Renvoie l’objet de plage associé à la ligne de total du tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableObject.getTotalRowRange();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[Range](range.md)

#### <a name="examples"></a>Exemples
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableTotalsRange = table.getTotalRowRange();
    tableTotalsRange.load('address');    
    return ctx.sync().then(function() {
            console.log(tableTotalsRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="reapplyfilters"></a>reapplyFilters()
Applique de nouveau tous les filtres actuellement appliqués sur le tableau.

#### <a name="syntax"></a>Syntaxe
```js
tableObject.reapplyFilters();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void
### <a name="property-access-examples"></a>Exemples d’accès aux propriétés

Obtenir un tableau par son nom 

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.load('index')
    return ctx.sync().then(function() {
            console.log(table.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Obtenir un tableau par son indice

```js
Excel.run(function (ctx) { 
    var index = 0;
    var table = ctx.workbook.tables.getItemAt(0);
    table.load('id')
    return ctx.sync().then(function() {
            console.log(table.id);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Définir le style du tableau 

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.name = 'Table1-Renamed';
    table.showTotals = false;
    table.style = 'TableStyleMedium2';
    table.load('tableStyle');
    return ctx.sync().then(function() {
            console.log(table.style);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="setting-formulas"></a>Définition de formules

#### <a name="common-pitfalls-when-setting-formulas-in-excel-from-add-ins"></a>Problèmes courants lors de la définition de formules dans Excel à partir de compléments

par Zlatko Michailov  
Microsoft Corp.


Cet article décrit trois problèmes que les développeurs de compléments Excel peuvent rencontrer, ainsi que les moyens de les contourner. Il est important de comprendre ces scénarios, en particulier car ils n’entraînent pas un échec dans des circonstances normales pour les compléments. Le complément peut paraître parfaitement normal lorsqu’il est utilisé sur une petite plage. Cependant, ses performances peuvent se dégrader de manière linéaire au fur et à mesure que sa plage de fonctionnement cible augmente au fil du temps.

Les deux premiers problèmes se manifestent lorsque les formules sont définies dans des colonnes du __tableau__, des colonnes spécifiques contenant des formules et des colonnes avec une ligne de totaux.

##### <a name="setting-formulas-in-calculated-table-columns"></a>Définition de formules dans des colonnes de tableau calculées

Cet [article](https://support.office.com/en-us/article/Use-calculated-columns-in-an-Excel-table-873FBAC6-7110-4300-8F6F-AAFA2EA11CE8) propose une vue d’ensemble des colonnes calculées.

La fonctionnalité clé est décrite à l’étape 4 :

> Lorsque vous appuyez sur Entrée, la formule est automatiquement appliquée aux cellules de la colonne, au-dessus et au-dessous de la cellule dans laquelle vous avez entré la formule. La formule est identique pour chaque ligne, mais comme il s’agit d’une référence structurée, Excel sait identifier chaque ligne en interne.

Cela signifie que chaque mise à jour de formule peut être multipliée N fois, N étant le nombre de lignes du tableau.

Les utilisateurs ne remarqueront peut-être aucun décalage significatif sur un tableau à 1 000 lignes, mais toute interaction avec un tableau contenant 10 000 lignes peut entraîner une expérience dégradée.

Heureusement, le calcul automatique des colonnes d’Excel est suffisamment intelligent et vous ne remarquerez peut-être pas le problème décrit ci-dessus. Pour qu’une colonne soit recalculée automatiquement, elle doit être vide ou entièrement calculée automatiquement. Si vous brisez la « pureté » de la colonne en insérant une valeur (et non une formule) dans une cellule, Excel ne tente pas de la recalculer automatiquement. En outre, si vous essayez de définir la formule qu’Excel a déjà définie dans cette colonne, le recalcul serait une absence d’opération.

Par exemple, supposons que vous définissez la formule `=B2+C2` sur la cellule `A2`. Si la colonne est vide, Excel calcule toutes les cellules de cette colonne _en ajustant l’index de ligne_. Ensuite, lorsque vous passez à la ligne suivante et que vous définissez la formule `=B3+C3` sur `A3`, aucun nouveau calcul de colonne n’est effectué, car cette formule est déjà définie automatiquement sur la colonne entière.

Toutefois, si vous souhaitez que la colonne représente une fonction de l’index de ligne, par exemple `=i * i` où _i_ est l’index de ligne, cela cause non seulement un recalcul de l’ensemble de la colonne à chaque mise à jour, mais vous obtenez également une colonne qui affiche la même (dernière) formule.

##### <a name="setting-formulas-on-a-table-with-a-totals-row"></a>Définition de formules dans un tableau avec une ligne de totaux

La définition de formules dans des tableaux avec une ligne de totaux activée peut parfois provoquer des problèmes de performances. Il est important d’indiquer que même une ligne de totaux par défaut, c'est-à-dire avec une valeur statique dans la cellule la plus à gauche et une valeur `Count` dans les cellules les plus à droite, et avec toutes ses cellules comprises entre les deux valeurs `None` pourrait reproduire le problème. 

Bien qu’il existe une solution de contournement plus simple (définir toutes les formules, puis ajouter la ligne de totaux dans le tableau), la solution générique recommandée pour les deux problèmes ci-dessus est l’utilisation d’une plage brute lors de la définition des formules, puis la conversion de cette plage en tableau.

Voici une fonction générique qui met à jour une plage de données et crée un tableau sur la plage cible. 

```js
function createAndPopulateTable(context, worksheetName, rangeAddress, hasHeaderRow, headerValues, bodyFormulas, tableCustomizer) {
    var worksheet = context.workbook.worksheets.getItem(worksheetName);

    // Calculate table-, body-, and header- ranges
    var tableRange = worksheet.getRange(rangeAddress);
    var bodyRange = tableRange;
    if (hasHeaderRow) {
        bodyRange = tableRange.getResizedRange(-1, 0).getOffsetRange(1, 0);
        if (headerValues) {
            // Set header values
            var headerRange = tableRange.getRow(0);
            headerRange.values = headerValues;
        }
    }
    
    // Set body formulas
    bodyRange.formulas = bodyFormulas;

    return context.sync()
        .then(function() {
            // Create the table
            var table = context.workbook.tables.add(tableRange, hasHeaderRow);

            // Invoke the caller's customizer
            if (tableCustomizer) {
                tableCustomizer(table);
            }

            return context.sync();
        });
}
```

La fonction ci-dessus est disponible en ligne à un [emplacement public](https://gist.github.com/zlatko-michailov/2b0418c986d9da6ee0bdf7aa346d3a4f).

Elle peut être utilisée comme suit :
```js
    return Excel.run(function(context) {
        return createAndPopulateTable(context, "Sheet1", "B3:E6", true, [['Alpha', 'Beta', 'Gamma', 'Delta']], 
                    [ ['=1+1', null, null, '=B4'], 
                      ['=2+2', null, null, '=B5'],
                      ['=3+3', null, null, '=B6'] ],
                    function (table) {
                        table.style = 'TableStyleLight1';
                        table.showTotals = true;
                    });
    });
```

Le calcul automatique des colonnes peut être désactivé dans le client de bureau Excel (il est activé par défaut), mais il est toujours activé dans Excel Online. Par conséquent, en tant que développeur de compléments, vous devez supposer qu’il est activé pour la majorité des utilisateurs de votre complément.


##### <a name="getting-a-range-object"></a>Obtention d'un objet Range

Ce problème est propre à l’implémentation de l’API JavaScript.

Pour un suivi correct de la plage lors des insertions et des suppressions de lignes/colonnes, une liaison est créée en interne chaque fois qu’un objet `Range` est demandé. Par la suite, lorsqu’une cellule est mise à jour, toutes les liaisons appropriées doivent être informées pour se mettre à jour.

Ainsi, le code suivant (ligne 8), qui semble anodin du point de vue de la programmation générale, augmente la complexité de façon exponentielle :
```js
    Excel.run(function(context) {
        var n = 10000;
        var worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.load();

        var arr = [];
        for (var i = 2; i <= n + 1; i++) {
            var range = worksheet.getRange("C3:C" + (n + 1)); /* <-- PROBLEM! */
            arr.push(["=A" + i + " + B" + i]);
        }
        range.formulas = arr; 
        return context.sync();
    });
```

La solution de contournement consiste à éviter tout recours inutile à l’objet `Range` en sortant la ligne pertinente de la boucle :
```js
    Excel.run(function(context) {
        var n = 10000;
        var worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.load();

        var arr = [];
        var range = worksheet.getRange("C3:C" + (n + 1)); /* <-- OK */
        for (var i = 2; i <= n + 1; i++) {
            arr.push(["=A" + i + " + B" + i]);
        }
        range.formulas = arr; 
        return context.sync();
    });
```
