# <a name="uidisplaydialogasync-method"></a>Méthode UI.displayDialogAsync

Affiche une boîte de dialogue dans un hôte Office. 

## <a name="requirements"></a>Configuration requise

|Hôte|Nouveauté de|Dernière modification dans |
|:---------------|:--------|:----------|
|Word, Excel, PowerPoint|1.1|1.1|
|Outlook|Mailbox 1.4|Mailbox 1.4|

Cette méthode est disponible dans l’[ensemble de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md) DialogApi pour les compléments Word, Excel ou PowerPoint, et dans l’ensemble de conditions requises Mailbox 1.4 pour Outlook. Pour spécifier l’ensemble de conditions requises DialogAPI, utilisez le code suivant dans votre manifeste.

```xml
<Requirements> 
  <Sets DefaultMinVersion="1.1"> 
    <Set Name="DialogApi"/> 
  </Sets> 
</Requirements> 
```

Pour spécifier l’ensemble de conditions requises Mailbox 1.4, utilisez le code suivant dans votre manifeste.

```xml
<Requirements> 
  <Sets DefaultMinVersion="1.4"> 
    <Set Name="Mailbox"/> 
  </Sets> 
</Requirements> 
```

Pour détecter cette API en cours d’exécution dans un complément Word, Excel ou PowerPoint, utilisez le code suivant.

```js
if (Office.context.requirements.isSetSupported('DialogApi', 1.1)) {  
  // Use Office UI methods; 
} else { 
  // Alternate path 
} 
```

Pour détecter cette API en cours d’exécution dans un complément Outlook, utilisez le code suivant.

```js
if (Office.context.requirements.isSetSupported('Mailbox', 1.4)) {  
  // Use Office UI methods; 
} else { 
  // Alternate path 
} 
```

Vous pouvez également vérifier si la méthode `displayDialogAsync` n’est pas définie avant de l’utiliser.

```js
if (Office.context.ui.displayDialogAsync !== undefined) {
  // Use Office UI methods
}
```

### <a name="supported-platforms"></a>Plateformes prises en charge
Pour plus d’informations sur les plateformes prises en charge, voir la page relative aux [ensembles de conditions requises DialogAPI](../requirement-sets/dialog-api-requirement-sets.md).

## <a name="syntax"></a>Syntaxe

```js
Office.context.ui.displayDialogAsync(startAddress, options, callback);
```
##<a name="examples"></a>Exemples

Pour obtenir un exemple simple qui utilise la méthode **displayDialogAsync**, consultez l’[exemple de boîte de dialogue API de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example/) sur GitHub.

Pour obtenir des exemples de scénario d’authentification, consultez les pages suivantes :

- [Complément PowerPoint dans Microsoft Graph - ASP.Net - Insérer un graphique](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Complément Office dans Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Complément Excel - ASP.NET - QuickBooks](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [Exemple d’authentification du complément Office sur le serveur pour ASP.net MVC](https://github.com/dougperkes/Office-Add-in-AspNetMvc-ServerAuth/tree/Office2016DisplayDialog)
- [Authentification client Office 365 du complément Office pour AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)


 
## <a name="parameters"></a>Paramètres

| Paramètre       | Type    |Description|
|:---------------|:--------|:----------|
|startAddress|string|Accepte l’URL HTTPS(TLS) initiale qui s’ouvre dans la boîte de dialogue. <ul><li>La page initiale doit figurer sur le même domaine que la page parent. Après le chargement de la page initiale, vous pouvez également accéder à d’autres domaines.</li><li>Toute page appelant [office.context.ui.messageParent](officeui.messageparent.md) doit également figurer sur le même domaine que la page parent.</li></ul>|
|options|object|Facultatif. Accepte un objet options pour définir les comportements de la boîte de dialogue.|
|callback|object|Accepte une méthode de rappel pour gérer la tentative de création de boîte de dialogue.|
    
### <a name="configuration-options"></a>Options de configuration
Les options de configuration suivantes sont disponibles pour une boîte de dialogue.


| Propriété       | Type    |Description|
|:---------------|:--------|:----------|
|**width**|int|Facultatif. Définit la largeur de la boîte de dialogue sous forme de pourcentage de l’affichage actuel. La valeur par défaut est 80 %. La résolution minimale est de 250 pixels.|
|**height**|int|Facultatif. Définit la hauteur de la boîte de dialogue sous forme de pourcentage de l’affichage actuel. La valeur par défaut est 80 %. La résolution minimale est de 150 pixels.|
|**displayInIframe**|bool|Facultatif. Détermine si la boîte de dialogue doit être affichée dans un IFrame. **Ce paramètre n’est applicable que dans les clients Office Online** et est ignoré par les autres plateformes. Les valeurs possibles sont les suivantes :<ul><li>False (valeur par défaut) : la boîte de dialogue s’affichera dans une nouvelle fenêtre de navigateur (fenêtre contextuelle). Recommandé pour les pages d’authentification qui ne peuvent pas être affichées dans un IFrame. </li><li>True : la boîte de dialogue s’affichera sous la forme d’une fenêtre flottante avec un IFrame. Recommandé pour une expérience utilisateur et des performances optimales.</li>|


## <a name="callback-value"></a>Valeur de rappel
Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **displayDialogAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Accéder à l’objet [Dialog](../../reference/shared/officeui.dialog.md).|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou objet défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|

### <a name="errors-from-displaydialogasync"></a>Erreurs provenant de displayDialogAsync

En plus des erreurs générales liées à la plateforme et au système, les erreurs suivantes sont propres à l’appel de la méthode **displayDialogAsync**.

|**Numéro de code**|**Signification**|
|:-----|:-----|
|12004|Le domaine de l’URL transmis à `displayDialogAsync` n’est pas approuvé. Le domaine doit soit être identique à la page hôte (y compris le protocole et le numéro de port), soit être inscrit dans la section `<AppDomains>` du manifeste du complément.|
|12005|L’URL transmise à `displayDialogAsync` utilise le protocole HTTP. C’est le protocole HTTPS qui est requis. (Dans certaines versions d’Office, le message d’erreur renvoyé avec le code 12005 est identique à celui renvoyé avec le code 12004.)|
|12007|Une boîte de dialogue est déjà ouverte à partir du volet Office. Une seule boîte de dialogue à la fois peut être ouverte dans un complément de volet Office.|



## <a name="design-considerations"></a>Considérations relatives à la conception
Les considérations relatives à la conception ci-dessous s’appliquent aux boîtes de dialogue :

- Un complément Office ne peut comporter qu’une seule boîte de dialogue ouverte à la fois.
- Toutes les boîtes de dialogue peuvent être déplacées et redimensionnées par l’utilisateur.
- Toutes les boîtes de dialogue s’affichent au centre de l’écran à l’ouverture.
- Les boîtes de dialogue s’affichent au-dessus de l’application hôte et dans l’ordre dans lequel elles ont été créées.

Utilisez une boîte de dialogue pour :

- Afficher les pages d’authentification permettant de collecter les informations d’identification de l’utilisateur.
- Afficher un écran d’erreur/de progression/de saisie à partir d’une commande ShowTaskpane ou ExecuteAction.
- Augmenter provisoirement la surface dont un utilisateur dispose pour effectuer une tâche.

N’utilisez pas de boîte de dialogue pour interagir avec un document. Il est préférable d’utiliser un volet des tâches. 

Pour obtenir un exemple de modèle de conception à utiliser pour créer une boîte de dialogue, consultez [Boîte de dialogue client](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Client_Dialog.md) dans le référentiel relatif aux modèles de conception de l’expérience utilisateur du complément Office sur GitHub.
