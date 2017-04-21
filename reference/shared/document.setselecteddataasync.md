
# <a name="documentsetselecteddataasync-method"></a>Méthode Document.setSelectedDataAsync
Écrit des données dans la sélection actuelle au sein du document.

|||
|:-----|:-----|
|**Hôtes :** Access, Excel, PowerPoint, Project, Word, Word Online|**Types de complément :  **Contenu, volet Office|
|**Disponible dans l’[ensemble de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Dernière modification dans**|1.1|

```js
Office.context.document.setSelectedDataAsync(data [, options], callback(asyncResult));
```


## <a name="parameters"></a>Paramètres

|Nom       | Type  | Description
|:----------|:------|:-----
| data      |objet | Les données peuvent être de n’importe quel [type de coercition](#coerciontype) pris en charge
| options   |objet | Spécifie un ensemble de [paramètres facultatifs](#options)
| callback  |objet | [AsyncResult](../../reference/shared/asyncresult.md), objet 



## <a name="options"></a>Options
```js
{
    coercionType: '',
    tableOptions: [],
    cellFormat: [],
    imageLeft: 0,
    imageTop: 0,
    imageWidth: 0,
    imageHeight: 0,
    asyncContext
}
```

### <a name="coerciontype"></a>coercionType
Les types de contrainte suivants sont pris en charge par Office.js. Notez que certains types de contrainte ne sont pas pris en charge par certains hôtes. 

|Nom                       |Access |Excel  |Word   |PowerPoint
|:--------------------------|:-----:|:-----:|:-----:|:---------:|
|Office.CoercionType.Text   |       |   X   |   X   |   X       |
|Office.CoercionType.Matrix |       |   X   |   X   |           |
|Office.CoercionType.Table  |   X   |   X   |   X   |           |
|Office.CoercionType.Html   |       |       |   X   |           |
|Office.CoercionType.Ooxml  |       |       |   X   |           |
|Office.CoercionType.Image  |       |   X   |   X   |   X       |

### <a name="tableoptions-object"></a>tableOptions (objet)
Pour le tableau inséré, liste de paires clé-valeur qui spécifient les options de mise en forme de tableau, comme la ligne d’en-tête, le nombre total de lignes et les lignes à bandes. (ajouté dans 1.1)

### <a name="cellformat-object"></a>CellFormat (objet)
Pour le tableau inséré, liste de paires clé-valeur qui spécifient la plage de cellules, lignes ou colonnes et la mise en forme de cellule à appliquer à cette plage. (ajouté dans 1.1)

### <a name="imageleft-number"></a>imageLeft (nombre)
Cette option s’applique à l’insertion des images. Indique l’emplacement d’insertion par rapport au côté gauche de la diapositive pour PowerPoint et sa relation avec la cellule actuellement sélectionnée dans Excel. Cette valeur est ignorée pour Word. Cette valeur est exprimée en points.

### <a name="imagetop-number"></a>imageTop (nombre)
Cette option s’applique à l’insertion des images. Indique l’emplacement d’insertion par rapport à la partie supérieure de la diapositive PowerPoint et sa relation avec la cellule actuellement sélectionnée dans Excel. Cette valeur est ignorée pour Word. Cette valeur est exprimée en points.

### <a name="imagewidth-number"></a>imageWidth (nombre)
Cette option s’applique à l’insertion des images. Indique la largeur de l’image. Si cette option est indiquée sans imageHeight, l’image sera dimensionnée pour correspondre à la valeur de la largeur de l’image. Si la largeur de l’image et la hauteur de l’image sont indiquées, l’image sera redimensionnée selon ces proportions. Si ni la hauteur ni la largeur de l’image est fournie, la taille de l’image par défaut et les proportions seront utilisées. Cette valeur est exprimée en points.

### <a name="imageheight-number"></a>imageHeight (nombre)
Cette option s’applique à l’insertion des images. Indique la hauteur de l’image. Si cette option est indiquée sans imageWidth, l’image sera dimensionnée pour correspondre à la valeur de la hauteur de l’image. Si la largeur de l’image et la hauteur de l’image sont indiquées, l’image sera redimensionnée selon ces proportions. Si ni la hauteur ni la largeur de l’image est fournie, la taille de l’image par défaut et les proportions seront utilisées. Cette valeur est exprimée en points.

### <a name="asynccontext-object--value"></a>asyncContext (objet | valeur)
Objet défini par l’utilisateur disponible sur la propriété asyncCesult de l’objet AsyncResult. Utilisez ce paramètre pour indiquer un objet ou une valeur à AsyncResult lorsque le rappel est une fonction nommée.


## <a name="callback-value"></a>Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **setSelectedDataAsync**, la propriété [AsyncResult.value](../../reference/shared/asyncresult.value.md) renvoie toujours **undefined**, car il n’existe aucun objet ni aucune donnée à récupérer.


## <a name="remarks"></a>Remarques

La valeur transmise pour le paramètre _data_ contient les données à écrire dans la sélection actuelle. Si la valeur est :


-  **Une chaîne :** Du texte brut ou tout élément dont le type peut être forcé en type **string** sera inséré.
    
    
    
    Dans Excel, vous pouvez également spécifier le paramètre _data_ en tant que formule valide pour ajouter cette dernière à la cellule sélectionnée. Par exemple, la définition du paramètre _data_ sur `"=SUM(A1:A5)"` totalisera les valeurs de la plage spécifiée. Toutefois, après avoir défini une formule sur la cellule liée, vous ne pouvez pas lire la formule ajoutée (ni les formules préexistantes) à partir de la cellule liée. Si vous appelez la méthode [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) sur la cellule sélectionnée pour en lire les données, la méthode peut renvoyer uniquement les données affichées dans la cellule (le résultat de la formule).
    
-  **Un tableau de tableaux (« matrice ») :** Des données tabulaires sans en-tête seront insérées. Par exemple, pour écrire des données sur trois lignes dans deux colonnes, vous pouvez transmettre un tableau comme suit : `[["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`. Pour écrire une seule colonne de trois lignes, transmettez un tableau comme suit : `[["R1C1"], ["R2C1"], ["R3C1"]]`
    
    
    
    Dans Excel, vous pouvez également spécifier le paramètre _data_ en tant que tableau de tableaux contenant des formules valides pour les ajouter aux cellules sélectionnées. Par exemple, si aucune autre donnée n’est remplacée, la définition du paramètre _data_ sur `[["=SUM(A1:A5)","=AVERAGE(A1:A5)"]]` ajoutera ces deux formules à la sélection. Comme lors de la définition d’une formule sur une cellule unique en tant que « texte », vous ne pouvez pas lire les formules ajoutées (ni les formules préexistantes) après leur définition. Vous pouvez uniquement lire les résultats des formules.
    
-  **Un objet [TableData](../../reference/shared/tabledata.md) :** Un tableau avec des en-têtes est inséré.
    
    
    
     >**Remarque :** dans Excel, si vous spécifiez des formules dans l’objet **TableData** que vous passez au paramètre _data_, vous risquez d’obtenir des résultats différents de ceux que vous attendez, en raison de la fonctionnalité d’Excel « Colonnes calculées », qui duplique automatiquement les formules dans une colonne. Pour contourner ce problème lorsque vous souhaitez écrire un paramètre _data_ contenant des formules pour une table sélectionnée, spécifiez les données sous forme de tableau de tableaux (au lieu de les spécifier sous forme d’objet **TableData**) et définissez le paramètre _coercionType_ sur **Microsoft.Office.Matrix** ou « matrice ».
    
### <a name="application-specific-behaviors"></a>Comportements propres à l’application

En outre, les actions suivantes (propres aux applications) s’appliquent lors de l’écriture de données dans une sélection.

#### <a name="word"></a>Word


- S’il n’y a aucune sélection et que le point d’insertion se trouve à un emplacement valide, le contenu du paramètre _data_ spécifié est inséré au point d’insertion comme suit :
    
      - Si le paramètre _data_ contient une chaîne, le texte spécifié est inséré.
    
  - Si le paramètre _data_ contient un tableau de tableaux (« matrice ») ou un objet **TableData**, un nouveau tableau Word est inséré.
    
  - Si le paramètre _data_ contient du code HTML, le code HTML spécifié est inséré.
    
     >**Important** :  Si le code HTML que vous insérez n’est pas valide, Word ne déclenche aucune erreur. Word insère autant de code HTML que possible et omet les données non valides.
  - Si le paramètre _data_ contient du code Office Open XML, le code XML spécifié est inséré.
    
  - Si le paramètre _data_ contient un flux d’images encodé en base64, l’image spécifiée est insérée.
    
- S’il existe une sélection, elle est remplacée par le contenu du paramètre _data_ spécifié selon les mêmes règles que ci-dessus.
    
-  **Insérer des images** : les images insérées sont placées en ligne. Les paramètres **imageLeft** et **imageTop** sont ignorés. Les proportions de l’image sont toujours verrouillées. Si seul un des paramètres **imageWidth** et **imageHeight** est donné, l’autre valeur est automatiquement redimensionnée pour conserver les proportions d’origine.
    
#### <a name="excel"></a>Excel


- Si une seule cellule est sélectionnée :
    
      - Si le paramètre _data_ contient une chaîne, le texte spécifié est inséré en tant que valeur de la cellule actuelle.
    
  - Si le contenu du paramètre _data_ est un tableau de tableaux (« matrice »), l’ensemble spécifié de lignes et de colonnes est inséré, à condition qu’aucune autre donnée des cellules environnantes ne soit remplacée.
    
  - Si le contenu du paramètre _data_ est un objet **TableData**, un nouveau tableau Excel avec l’ensemble spécifié de lignes et d’en-têtes est inséré, à condition qu’aucune autre donnée des cellules environnantes ne soit remplacée.
    
- Si plusieurs cellules sont sélectionnées et que la forme ne correspond pas à la forme du contenu du paramètre _data_, une erreur est renvoyée.
    
- Si plusieurs cellules sont sélectionnées et que la forme de la sélection correspond exactement à la forme du contenu du paramètre _data_, les valeurs des cellules sélectionnées sont mises à jour en fonction des valeurs du paramètre _data_.
    
-  **Insérer des images** : les images insérées sont flottantes. Les paramètres **imageLeft** et **imageTop** de position sont relatifs à la ou aux cellule(s) actuellement sélectionnée(s). Les valeurs **imageLeft** et **imageTop** négatives sont autorisées et éventuellement réajustées par Excel pour positionner l’image dans une feuille de calcul. Les proportions sont verrouillées à moins que les paramètres **imageWidth** et **imageHeight** soient tous deux indiqués. Si seul un des paramètres **imageWidth** et **imageHeight** est donné, l’autre valeur est automatiquement redimensionnée pour conserver les proportions d’origine.
    
Dans tous les autres cas, une erreur est renvoyée.

#### <a name="excel-online"></a>Excel Online

En plus des comportements décrits pour Excel ci-dessus, les limites suivantes s’appliquent lors de l’écriture de données dans Excel Online. 


- Le nombre total de cellules que vous pouvez écrire dans une feuille de calcul avec le paramètre _data_ ne peut pas dépasser 20 000 dans un appel unique à cette méthode.
    
- Le nombre de _groupes de mise en forme_ transmis au paramètre _cellFormat_ ne peut pas dépasser 100. Un groupe de mise en forme se compose d’un ensemble de mises en forme appliquées à une plage de cellules donnée. Par exemple, l’appel suivant transmet deux groupes de mise en forme au paramètre _cellFormat_.
    

```js
  Office.context.document.setSelectedDataAsync(
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});
```

#### <a name="powerpoint"></a>PowerPoint

Les images insérées sont flottantes. Les paramètres de position **imageLeft** et **imageTop** sont facultatifs, mais, s’ils sont indiqués, les deux doivent être présents. Si une seule valeur est indiquée, elle sera ignorée. Les valeurs **imageLeft** et **imageTop** négatives sont autorisées et peuvent positionner une image en dehors d’une diapositive. Si aucun paramètre facultatif n’est indiqué et qu’une diapositive présente un espace réservé, l’image remplacera l’espace réservé dans la diapositive. Les proportions de l’image seront verrouillées, sauf si les paramètres **imageWidth** et **imageHeight** sont tous deux indiqués. Si seul un des paramètres **imageWidth** et **imageHeight** est donné, l’autre valeur est automatiquement redimensionnée pour conserver les proportions d’origine.


## <a name="example"></a>Exemple

L’exemple suivant affecte à la cellule ou au texte sélectionné la valeur « Hello World! ». En cas d’échec, la valeur de la propriété [error.message](../../reference/shared/error.message.md) est affichée.


```js
function writeText() {
    Office.context.document.setSelectedDataAsync("Hello World!",
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                 write(error.name + ": " + error.message);
            }
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



En spécifiant le paramètre facultatif _coercionType_, vous pouvez indiquer le type de données que vous souhaitez écrire dans une sélection. L’exemple suivant écrit des données sous forme d’un tableau de deux colonnes et trois lignes, en spécifiant _coercionType_ en tant que `"matrix"` pour cette structure de données. En cas d’échec, la valeur de la propriété [error.message](../../reference/shared/error.message.md) est affichée.




```js
function writeMatrix() {
    Office.context.document.setSelectedDataAsync([["Red", "Rojo"], ["Green", "Verde"], ["Blue", "Azul"]], {coercionType: Office.CoercionType.Matrix}
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                write(error.name + ": " + error.message);
            }
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



L’exemple suivant écrit des données sous forme d’un tableau d’une seule colonne avec un en-tête et quatre lignes, en spécifiant _coercionType_ en tant que `"table"` pour cette structure de données. En cas d’échec, la valeur de la propriété [error.message](../../reference/shared/error.message.md) est affichée.




```js
function writeTable() {
    // Build table.
    var myTable = new Office.TableData();
    myTable.headers = [["Cities"]];
    myTable.rows = [['Berlin'], ['Roma'], ['Tokyo'], ['Seattle']];

    // Write table.
    Office.context.document.setSelectedDataAsync(myTable, {coercionType: Office.CoercionType.Table},
        function (result) {
            var error = result.error
            if (result.status === Office.AsyncResultStatus.Failed) {
                write(error.name + ": " + error.message);
            }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



 Dans Word, si vous voulez écrire du contenu HTML dans la sélection, vous pouvez spécifier le paramètre _coercionType_ en tant que `"html"` comme indiqué dans l’exemple suivant. Ce dernier utilise les balises HTML `<b>` pour mettre la chaîne « Hello » en gras.




```js
function writeHtmlData() {
    Office.context.document.setSelectedDataAsync("<b>Hello</b> World!", {coercionType: Office.CoercionType.Html}, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            write('Error: ' + asyncResult.error.message);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Dans Word, PowerPoint ou Excel, si vous souhaitez écrire une image dans la sélection, vous pouvez spécifier le paramètre _coercionType_ en tant que `"image"`, comme illustré dans l’exemple suivant. Notez qu’imageLeft et imageTop sont ignorées par Word.




```js
function insertPictureAtSelection(base64EncodedImageStr) {

    Office.context.document.setSelectedDataAsync(base64EncodedImageStr, {
       coercionType: Office.CoercionType.Image,
       imageLeft: 50,
       imageTop: 50,
       imageWidth: 100,
       imageHeight: 100
       },
       function (asyncResult) {
           if (asyncResult.status === Office.AsyncResultStatus.Failed) {
               console.log("Action failed with error: " + asyncResult.error.message);
           }
       });
}
```


## <a name="support-details"></a>Informations de prise en charge


Une coche (![Symbole de coche](../../images/mod_off15_checkmark.png)) dans la matrice suivante indique que cette méthode est prise en charge dans l’application hôte Office correspondante. Une cellule vide indique que l’application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**

||**Office pour bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**|![Symbole de coche](../../images/mod_off15_checkmark.png)|||
|**Excel**|![Symbole de coche](../../images/mod_off15_checkmark.png)|![symbole de coche](../../images/mod_off15_checkmark.png)|![Symbole de coche](../../images/mod_off15_checkmark.png)|
|**PowerPoint**|![Symbole de coche](../../images/mod_off15_checkmark.png)|![symbole de coche](../../images/mod_off15_checkmark.png)|![Symbole de coche](../../images/mod_off15_checkmark.png)|
|**Word**|![Symbole de coche](../../images/mod_off15_checkmark.png)|![symbole de coche](../../images/mod_off15_checkmark.png)|![Symbole de coche](../../images/mod_off15_checkmark.png)|


|||
|:-----|:-----|
|**Disponible dans les ensembles de conditions requises**|Selection|
|**Niveau d’autorisation minimal**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## <a name="support-notes"></a>Notes de prise en charge
**Modifié dans :** 1.1. La prise en charge des composants de contenu pour Access exige l’ensemble de conditions requises **Selection** 1.1 ou ultérieur. La prise en charge de la définition des données d’image nécessite l’ensemble de conditions requises **ImageCoercion** 1.1 ou ultérieur. Pour définir l’activation de l’application, utilisez le code suivant :

```xml
<Requirements>
    <Sets DefaultMinVersion="1.1">
        <Set Name="ImageCoercion"/>
    </Sets>
</Requirements>
```

La détection d’exécution de la fonctionnalité ImageCoercion peut être effectuée par le code suivant :

```javascript
if (Office.context.requirements.isSetSupported('ImageCoercion', '1.1')) {)) {
    // insertViaImageCoercion();
} 
else {
    // insertViaOoxml();
}
```

## <a name="support-history"></a>Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Dans Word et Word Online, l’écriture de données sous la forme d’un flux d’images codées en base64 est désormais prise en charge.|
|1.1|Dans Word Online, l’écriture de _données_ en tant que **tableau** de tableaux (matrice) et **TableData** (tableau) est désormais prise en charge.|
|1.1|Dans Excel, PowerPoint et Word dans Office pour iPad, le même niveau de prise en charge que dans Excel, PowerPoint et Word sur le bureau Windows est désormais pris en charge.|
|1.1|Dans Word Online, l’écriture de _données_ en tant que **chaîne** (texte) est désormais prise en charge.|
|1.1|Prise en charge supplémentaire de la [définition de la mise en forme lors de l’insertion de tableaux](../../docs/excel/format-tables-in-add-ins-for-excel.md) avec des compléments pour Excel à l’aide des paramètres facultatifs _tableOptions_ et _cellFormat_.|
|1.1|Prise en charge supplémentaire de l’écriture de données de tableau dans les compléments pour Access.|
|1.0|Introduit|
