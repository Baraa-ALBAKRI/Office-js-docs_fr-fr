
# <a name="context.officetheme-property"></a>Propriété Context.officeTheme
Permet d’accéder aux propriétés pour les couleurs du thème Office.

 **Important :** Cette API fonctionne actuellement uniquement dans Excel, Outlook, PowerPoint et Word dans [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) sur le bureau Windows.


|||
|:-----|:-----|
|**Hôtes :**|Excel, Outlook, PowerPoint, Word|
|**Disponible dans l’[ensemble de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Pas dans un ensemble|
|**Ajouté dans**|1.3|



```js
Office.context.officeTheme
```


## <a name="members"></a>Membres


**Propriétés**

|||
|:-----|:-----|
|Nom|Description|
|[bodyBackgroundColor ](../../reference/shared/office.context.bodybackgroundcolor.md)|Obtient la couleur d’arrière-plan du corps du thème Office.|
|[bodyForegroundColor](../../reference/shared/office.context.bodyforegroundcolor.md)|Obtient la couleur de premier plan du corps du thème Office.|
|[controlBackgroundColor](../../reference/shared/office.context.controlbackgroundcolor.md)|Obtient la couleur d’arrière-plan du contrôle du thème Office.|
|[controlForegroundColor](../../reference/shared/office.context.controlforegroundcolor.md)|Obtient la couleur de premier plan du contrôle du thème Office.|

## <a name="remarks"></a>Remarques

À l’aide des couleurs du thème Office, vous pouvez coordonner le modèle de couleurs de votre complément avec le thème Office actuel sélectionné par l’utilisateur dans **Fichier**  >  **Compte Office**  >  **Thème Office**, qui est appliqué à toutes les applications hôtes Office. L’utilisation des couleurs du thème Office est appropriée pour les compléments Outlook et du volet Office.


## <a name="example"></a>Exemple


```js
function applyOfficeTheme(){
    // Get office theme colors.
    var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
    var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
    var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
    var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

    // Apply body background color to a CSS class.
    $('.body').css('background-color', bodyBackgroundColor);
}
```


## <a name="support-details"></a>Informations de prise en charge



|||
|:-----|:-----|
|**Niveau d’autorisation minimal**|[Restreint](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|De contenu, de volet de tâche, Outlook|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Office|

## <a name="support-history"></a>Historique de prise en charge


|**Version**|**Modifications**|
|:-----|:-----|
|1.3|Introduit|
