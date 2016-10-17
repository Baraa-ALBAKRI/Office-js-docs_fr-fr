# <a name="inlinepicture-object-(javascript-api-for-word)"></a>Objet InlinePicture (interface API JavaScript pour Word)

Représente une image incluse.

_S’applique à : Word 2016, Word pour iPad, Word pour Mac, Word Online_

## <a name="properties"></a>Propriétés
| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|altTextDescription|string|Obtient ou définit une chaîne qui représente le texte de remplacement associé à l’image incluse.|
|altTextTitle|string|Obtient ou définit une chaîne qui contient le titre de l’image incluse.|
|hyperlink|string|Obtient ou définit le lien hypertexte associé à l’image incluse.|
|lockAspectRatio|bool|Obtient ou définit une valeur qui indique si l’image incluse conserve ses proportions d’origine lorsque vous la redimensionnez.|

## <a name="relationships"></a>Relations
| Relation | Type   |Description|
|:---------------|:--------|:----------|
|height|**float**|Obtient ou définit un nombre qui décrit la hauteur de l’image incluse. Ceci est exprimé en points. |
|parentContentControl|[ContentControl](contentcontrol.md)|Obtient le contrôle de contenu qui contient l’image incluse. Renvoie null s’il n’existe pas de contrôle de contenu parent. En lecture seule.|
|paragraph|[paragraph](paragraph.md)|Obtient le paragraphe qui contient l’image incluse. En lecture seule.
|width|**float**|Obtient ou définit un nombre qui décrit la largeur de l’image incluse. Ceci est exprimé en points.|

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Supprime l’image du document.|
|[getBase64ImageSrc()](#getbase64imagesrc)|objet|Obtient un objet dont la valeur est la représentation de chaîne encodée au format Base64 de l’image incluse.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Insère un saut à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Encadre l’image incluse avec un contrôle de contenu de texte enrichi.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Insère un document dans le corps à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Insère du code HTML à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Insère une image dans le corps à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Replace » (remplacer), « Before » (avant) ou « After » (après). |
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Insère du code OOXML à l’emplacement spécifié.  La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Insère un paragraphe à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Insère du texte dans le corps à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Sélectionne l’image et y accède via l’interface utilisateur de Word. Les valeurs selectionMode peuvent être « Select » (sélectionner), « Start » (début) ou « End » (fin).|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## <a name="method-details"></a>Détails de méthodes

### <a name="delete()"></a>delete()
Supprime l’image du document.

#### <a name="syntax"></a>Syntaxe
```js
inlinePictureObject.delete();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
void

### <a name="getbase64imagesrc()"></a>getBase64ImageSrc()
Obtient un objet dont la valeur est la représentation de chaîne encodée au format Base64 de l’image incluse.

#### <a name="syntax"></a>Syntaxe
```js
var base64 = inlinePictureObject.getBase64ImageSrc();
return context.sync().then(function () {    
    console.log("base64 string is " + base64.value);
});

```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
object 



### <a name="insertbreak(breaktype:-breaktype,-insertlocation:-insertlocation)"></a>insertBreak(breakType: BreakType, insertLocation: InsertLocation)

#### <a name="syntax"></a>Syntaxe
```js
inlinePictureObject.insertBreak(breakType, insertLocation);
```
#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|breakType|BreakType|Obligatoire. Type de saut à ajouter au corps.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Before » (avant) » ou « After » (après).|

#### <a name="returns"></a>Retourne
void

### <a name="insertcontentcontrol()"></a>insertContentControl()
Encadre l’image incluse avec un contrôle de contenu de texte enrichi.

#### <a name="syntax"></a>Syntaxe
```js
inlinePictureObject.insertContentControl();
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
[ContentControl](contentcontrol.md)

### <a name="insertfilefrombase64(base64file:-string,-insertlocation:-insertlocation)"></a>insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Insère un document dans le corps à l’emplacement spécifié. La valeur insertLocation peut être « Before » (avant) ou « After » (après).

#### <a name="syntax"></a>Syntaxe
```js
inlinePictureObject.insertFileFromBase64(base64File, insertLocation);
```
#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|base64File|string|Obligatoire. Contenu d’un fichier docx encodé au format Base64.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Before » (avant) » ou « After » (après).|

#### <a name="returns"></a>Retourne
[Range](range.md)

### <a name="inserthtml(html:-string,-insertlocation:-insertlocation)"></a>insertHtml(html: string, insertLocation: InsertLocation)
Insère du code HTML à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).

#### <a name="syntax"></a>Syntaxe
```js
inlinePictureObject.insertHtml(html, insertLocation);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|Html|string|Obligatoire. Code HTML à insérer dans le document.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Before » (avant) » ou « After » (après).|

#### <a name="returns"></a>Retourne
[Range](range.md)


### <a name="insertinlinepicturefrombase64(base64encodedimage:-string,-insertlocation:-insertlocation)"></a>insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Insère une image dans le corps à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).

#### <a name="syntax"></a>Syntaxe
inlinePictureObject.insertInlinePictureFromBase64(image, insertLocation);

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Obligatoire. Image encodée au format Base64 à insérer dans le corps.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Before » (avant) » ou « After » (après).|

#### <a name="returns"></a>Retourne
[InlinePicture](inlinepicture.md)


### <a name="insertooxml(ooxml:-string,-insertlocation:-insertlocation)"></a>insertOoxml(ooxml: string, insertLocation: InsertLocation)
Insère du code OOXML à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).

#### <a name="syntax"></a>Syntaxe
```js
inlinePictureObject.insertOoxml(ooxml, insertLocation);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|ooxml|string|Obligatoire. Code OOXML à insérer.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Before » (avant) » ou « After » (après).|

#### <a name="returns"></a>Retourne
[Range](range.md)

### <a name="insertparagraph(paragraphtext:-string,-insertlocation:-insertlocation)"></a>insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Insère un paragraphe à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).

#### <a name="syntax"></a>Syntaxe
```js
inlinePictureObject.insertParagraph(paragraphText, insertLocation);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|paragraphText|string|Obligatoire. Texte de paragraphe à insérer.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Before » (avant) » ou « After » (après).|

#### <a name="returns"></a>Retourne
[Paragraph](paragraph.md)

### <a name="inserttext(text:-string,-insertlocation:-insertlocation)"></a>insertText(text: string, insertLocation: InsertLocation)
Insère du texte dans le corps à l’emplacement spécifié. La valeur insertLocation peut être « Before » (avant) ou « After » (après).

#### <a name="syntax"></a>Syntaxe
```js
inlinePictureObject.insertText(text, insertLocation);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|text|string|Obligatoire. Texte à insérer.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Before » (avant) » ou « After » (après).|

#### <a name="returns"></a>Retourne
[Range](range.md)

### <a name="select(selectionmode:-selectionmode)"></a>select(selectionMode: SelectionMode)
Sélectionne l’image et y accède via l’interface utilisateur de Word. Les valeurs selectionMode peuvent être « Select » (sélectionner), « Start » (début) ou « End » (fin).

#### <a name="syntax"></a>Syntaxe
```js
inlinePictureObject.select(selectionMode);
```

#### <a name="parameters"></a>Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Facultatif. Le mode de sélection peut être « Select » (sélectionner), « Start » (début) ou « End » (fin). « Select » (sélectionner) est la valeur par défaut.|

#### <a name="returns"></a>Retourne
void

### <a name="load(param:-object)"></a>load(param: object)
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

## <a name="support-details"></a>Informations de prise en charge
Utilisez l’[ensemble de conditions requises](../office-add-in-requirement-sets.md) dans les vérifications à l’exécution pour vous assurer que votre application est prise en charge par la version d’hôte de Word. Pour plus d’informations sur la configuration requise pour le serveur et l’application d’hôte Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).
