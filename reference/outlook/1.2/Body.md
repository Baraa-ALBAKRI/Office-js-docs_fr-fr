

# <a name="body"></a>Corps

L’objet `body` fournit des méthodes d’ajout et de mise à jour du contenu du message ou du rendez-vous. Il est renvoyé dans la propriété `body` de l’élément sélectionné.

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

### <a name="methods"></a>Méthodes

####  <a name="gettypeasync([options],-[callback])"></a>getTypeAsync([options], [callback])

Obtient une valeur qui indique si le contenu est au format HTML ou texte.

##### <a name="parameters:"></a>Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult).

Le type de contenu est renvoyé sous la forme d’une des valeurs [CoercionType](Office.md#coerciontype-string) dans la propriété `asyncResult.value`.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition|
####  <a name="prependasync(data,-[options],-[callback])"></a>prependAsync(data, [options], [callback])

Ajoute le contenu spécifié au début du corps de l’élément.

La méthode `prependAsync` insère la chaîne spécifiée au début du corps de l’élément. Appeler la méthode `prependAsync` revient à appeler la méthode [`setSelectedDataAsync`](#setselecteddataasyncdata-options-callback) avec le point d’insertion au début du contenu du corps.

Lorsque vous incluez des liens dans la balise HTML, vous pouvez désactiver l’aperçu du lien en ligne en définissant l’attribut `id` de l’ancre (`<a>`) sur `LPNoLP`. Par exemple :

```
Office.context.mailbox.item.body.prependAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### <a name="parameters:"></a>Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`data`| String||La chaîne doit être insérée au début du corps. Elle est limitée à un million de caractères.|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;optional&gt;</td><td>Format du corps souhaité. La chaîne du paramètre <code>data</code> est convertie dans ce format.</td></tr><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult). <br/>Les erreurs rencontrées seront indiquées dans la propriété `asyncResult.error`.<br/><table class="nested-table"><thead><tr><th>Code d'erreur</th><th>Description</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td>Le paramètre <code>data</code> comprend plus de 1 000 000 de caractères.</td></tr></tbody></table>|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Mode Outlook applicable| Composition|
####  <a name="setselecteddataasync(data,-[options],-[callback])"></a>setSelectedDataAsync(data, [options], [callback])

Remplace la sélection dans le corps par le texte spécifié.

La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps de l’élément ou, si du texte est sélectionné dans l’éditeur, elle remplace ce texte. Si le curseur ne s’est jamais trouvé dans le corps de l’élément, ou si le corps de l’élément n’est plus la partie active de l’interface utilisateur, la chaîne est insérée au début du corps de l’élément.

Lorsque vous incluez des liens dans la balise HTML, vous pouvez désactiver l’aperçu du lien en ligne en définissant l’attribut `id` de l’ancre (`<a>`) sur `LPNoLP`. Par exemple :

```
Office.context.mailbox.item.body.setSelectedDataAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### <a name="parameters:"></a>Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`data`| String||Chaîne à insérer dans le corps. Elle est limitée à un million de caractères.|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;optional&gt;</td><td>Format du corps souhaité. La chaîne du paramètre <code>data</code> est convertie dans ce format.</td></tr><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult). <br/>Les erreurs rencontrées seront indiquées dans la propriété `asyncResult.error`.<br/><table class="nested-table"><thead><tr><th>Code d'erreur</th><th>Description</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td>Le paramètre <code>data</code> comprend plus de 1 000 000 de caractères.</td></tr><tr><td><code>InvalidFormatError</code></td><td>Le type de corps est défini en mode HTML et le paramètre de données contient du texte brut.</td></tr></tbody></table>|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Mode Outlook applicable| Composition|
