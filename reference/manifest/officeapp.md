
# <a name="officeapp-element"></a>OfficeApp, élément
Élément racine dans le manifeste d’un complément Office.

 **Type de complément :** Application de contenu, de volet Office, de messagerie


## <a name="syntax:"></a>Syntaxe :


```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```


## <a name="contained-in:"></a>Contenu dans :

 _Aucun_


## <a name="must-contain:"></a>Doit contenir :



|**Élément**|**Contenu**|**Messagerie**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Id](../../reference/manifest/id.md)|x|x|x|
|[Version](../../reference/manifest/version.md)|x|x|x|
|[ProviderName](../../reference/manifest/providername.md)|x|x|x|
|[DefaultLocale](../../reference/manifest/defaultlocale.md)|x|x|x|
|[DefaultSettings](../../reference/manifest/defaultsettings.md)|x|x|x|
|[DisplayName](../../reference/manifest/displayname.md)|x|x|x|
|[Description](../../reference/manifest/description.md)|x|x|x|
|[FormSettings](../../reference/manifest/formsettings.md)||x||
|[Permissions](../../reference/manifest/permissions.md)|x||x|
|[Rule](../../reference/manifest/rule.md)||x||

## <a name="can-contain:"></a>Peut contenir :



|**Élément**|**Contenu**|**Messagerie**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[AlternateId](../../reference/manifest/alternateid.md)|x|x|x|
|[IconUrl](../../reference/manifest/iconurl.md)|x|x|x|
|[HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md)|x|x|x|
|[SupportUrl](../../reference/manifest/supporturl.md)|x|x|x|
|[AppDomains](../../reference/manifest/appdomains.md)|x|x|x|
|[Hosts](../../reference/manifest/hosts.md)|x|x|x|
|[Requirements](../../reference/manifest/requirements.md)|x|x|x|
|[AllowSnapshot](../../reference/manifest/allowsnapshot.md)|x|||
|[Permissions](../../reference/manifest/permissions.md)||x||
|[DisableEntityHighlighting](../../reference/manifest/disableentityhighlighting.md)||x||
|[Dictionary](../../reference/manifest/dictionary.md)|||x|
|[VersionOverrides](../../reference/manifest/versionoverrides.md)|X|X|X|

## <a name="attributes"></a>Attributs


|||
|:-----|:-----|
|xmlns|Définit la version de schéma et l’espace de noms du manifeste de complément Office. Cet attribut doit toujours être défini sur `"http://schemas.microsoft.com/office/appforoffice/1.1"`.|
|xmlns:xsi|Définit l’instance XMLSchema. Cet attribut doit toujours être défini sur `"http://www.w3.org/2001/XMLSchema-instance"`.|
|xsi:type|Définit le type de complément Office. Cet attribut doit être défini sur l’une des options suivantes : `"ContentApp"`, `"MailApp"` ou `"TaskPaneApp"`.|
