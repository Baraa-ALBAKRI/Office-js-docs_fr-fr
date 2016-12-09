# <a name="searchoptions-object-javascript-api-for-word"></a>Objet SearchOptions (interface API JavaScript pour Word)

Spécifie les options à inclure dans une opération de recherche.

_S’applique à : Word 2016, Word pour iPad, Word pour Mac, Word Online_

## <a name="properties"></a>Propriétés
| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|ignorePunct|bool|Obtient ou définit une valeur indiquant si toutes les marques de ponctuation entre les mots doivent être ignorées. Correspond à la case à cocher Ignorer les marques de ponctuation de la boîte de dialogue Rechercher et remplacer.|
|ignoreSpace|bool|Obtient ou définit une valeur indiquant si tous les espaces entre les mots doivent être ignorés. Correspond à la case à cocher Ignorer les caractères d’espacement de la boîte de dialogue Rechercher et remplacer.|
|matchCase|bool|Obtient ou définit une valeur indiquant si la recherche respecte la casse. Correspond à la case à cocher Respecter la casse de la boîte de dialogue Rechercher et remplacer (menu Édition).|
|matchPrefix|bool|Obtient ou définit une valeur indiquant si la recherche doit porter sur les mots qui commencent par la chaîne entrée. Correspond à la case à cocher Préfixe de la boîte de dialogue Rechercher et remplacer.|
|matchSoundsLike|bool|**Cette option a été déconseillée dans la mise à jour de juin 2016**. Obtient ou définit une valeur indiquant si la recherche doit porter sur les mots dont la prononciation est semblable à celle de la chaîne de recherche. Correspond à la case à cocher Recherche phonétique de la boîte de dialogue Rechercher et remplacer|
|matchSuffix|bool|Obtient ou définit une valeur indiquant si la recherche doit porter sur les mots qui se terminent par la chaîne entrée. Correspond à la case à cocher Suffixe de la boîte de dialogue Rechercher et remplacer.|
|matchWholeWord|bool|Obtient ou définit une valeur indiquant si la recherche doit uniquement porter sur des mots entiers et exclure le texte s’il est inclus dans un mot plus long. Correspond à la case à cocher Mot entier de la boîte de dialogue Rechercher et remplacer.|
|matchWildCards|bool|Obtient ou définit une valeur indiquant si la recherche est effectuée à l’aide d’opérateurs de recherche spéciaux. Correspond à la case Caractères génériques de la boîte de dialogue Rechercher et remplacer. Pour obtenir des informations importantes sur l’utilisation de cette option, consultez les conseils relatifs aux caractères génériques ci-dessous.|

_Voir des [exemples d’accès aux propriétés.](#property-access-examples)_

Les options de recherche sont facultatives. Elles doivent être définies à l’aide d’un littéral d’objet dans toutes les méthodes de recherche :

```js
    search('searchstring', {searchOption1:bool, ...searchOptionN:bool}
```

Vous pouvez fournir une ou plusieurs propriétés d’options de recherche dans le littéral d’objet pour définir les options de recherche. 

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## <a name="method-details"></a>Détails de méthodes

### <a name="loadparam-object"></a>load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### <a name="syntax"></a>Syntaxe
```js
object.load(param);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### <a name="returns"></a>Retourne
void

## <a name="property-access-examples"></a>Exemples d’accès aux propriétés

### <a name="ignore-punctuation-search"></a>Ignorer les signes de ponctuation dans la recherche
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document and ignore punctuation.
    var searchResults = context.document.body.search('video you', {ignorePunct: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-based-on-a-prefix"></a>Effectuer une recherche de préfixe
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document based on a prefix.
    var searchResults = context.document.body.search('vid', {matchPrefix: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-based-on-a-suffix"></a>Effectuer une recherche de suffixe
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to search the document for any string of characters after 'ly'.
    var searchResults = context.document.body.search('ly', {matchSuffix: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'orange';
            searchResults.items[i].font.highlightColor = 'black';
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-using-a-wildcard"></a>Effectuer une recherche à l’aide d’un caractère générique
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document with a wildcard
    // for any string of characters that starts with 'to' and ends with 'n'.
    var searchResults = context.document.body.search('to*n', {matchWildCards: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = 'pink';
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


## <a name="wildcard-guidance"></a>Aide concernant les caractères génériques 

| Pour trouver :         | Caractère générique |  Exemple |
|:-----------------|:--------|:----------|
| Un seul caractère| ? |s?t trouve sot et set. |
|Une chaîne de caractères| * |s*n son et solution.|
|Début d’un mot|< |<(intér) trouve intéressant et intérieur, mais pas désintéressé.|
|Fin d’un mot |> |(in)> trouve fin et besoin, mais pas origine.|
|Un des caractères spécifiés|[ ] |l[ea]s trouve les et las.|
|Tout caractère de cette plage| [-] |[b-d]arder trouve barder, carder et darder. Les plages doivent être définies dans l’ordre alphabétique ou croissant.|
|Tout caractère à l’exception de ceux de la plage entre les crochets|[!x-z] |p[!a-m]re trouve pore et pure, mais pas pare et pire.|
|Exactement n occurrences de l’expression ou du caractère précédent|n |bal\{2\}ade trouve ballade mais pas balade.|
|Au moins n occurrences de l’expression ou du caractère précédent|{n,} |bal{1,}ade recherche balade et ballade.|
|Entre n et m occurrences de l’expression ou du caractère précédent|{n,m} |10{1,3} trouve 10, 100 et 1000.|
|Une ou plusieurs occurrences de l’expression ou du caractère précédent|@ |mar@e trouve mare et marre.|

### <a name="escaping-the-special-characters"></a>Échappement des caractères spéciaux

La recherche avec des caractères génériques est essentiellement la même que la recherche sur une expression régulière. Il existe des caractères spéciaux dans les expressions régulières, notamment « [ », « ] », « ( »,« ) », « { », « } », « \* », « ? », « < », « > », « ! » et « @ ». Si l’un de ces caractères fait partie de la chaîne littérale que recherche le code, il doit être échappé, afin que Word sache qu’il faut le traiter littéralement et non dans le cadre de la logique de l’expression régulière. Pour échapper un caractère dans la fonction de recherche de l’interface utilisateur de Word, faites-le précéder d’un « \' », mais pour un échappement par programme, placez-le entre les caractères « [] ». Par exemple, « [\*]\* » recherche une chaîne qui commence par « \* », suivie d’autres caractères. 

## <a name="support-details"></a>Informations de prise en charge
Utilisez l’[ensemble de conditions requises](../office-add-in-requirement-sets.md) dans les vérifications à l’exécution pour vous assurer que votre application est prise en charge par la version d’hôte de Word. Pour plus d’informations sur la configuration requise pour le serveur et l’application d’hôte Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).
