# <a name="sectioncollection-object-(javascript-api-for-word)"></a>Objet SectionCollection (interface API JavaScript pour Word)

Contient la collection d’objets de [section](section.md)du document.

_S’applique à : Word 2016, Word pour iPad, Word pour Mac, Word Online_

## <a name="properties"></a>Propriétés
| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|items|[Section[]](section.md)|Collection d’objets de section. En lecture seule.|

## <a name="relationships"></a>Relations
Aucun


## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## <a name="method-details"></a>Détails de méthodes

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

#### <a name="examples"></a>Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy sectionsCollection object.
    var mySections = context.document.sections;

    // Queue a commmand to load the sections.
    context.load(mySections, 'body/style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Create a proxy object the primary header of the first section.
        // Note that the header is a body object.
        var myHeader = mySections.items[0].getHeader("primary");

        // Queue a command to insert text at the end of the header.
        myHeader.insertText("This is a header.", Word.InsertLocation.end);

        // Queue a command to wrap the header in a content control.
        myHeader.insertContentControl();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Added a header to the first section.");
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

## <a name="support-details"></a>Informations de prise en charge
Utilisez l’[ensemble de conditions requises](../office-add-in-requirement-sets.md) dans les vérifications à l’exécution pour vous assurer que votre application est prise en charge par la version d’hôte de Word. Pour plus d’informations sur la configuration requise pour le serveur et l’application d’hôte Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).
