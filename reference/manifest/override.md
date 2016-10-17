
# <a name="override-element"></a>Élément Override
Fournit une manière de spécifier la valeur d’un paramètre pour d’autres paramètres régionaux.

 **Type de complément :** Application de contenu, de volet Office, de messagerie


## <a name="syntax:"></a>Syntaxe :


```XML
<Override Locale="string " Value="string " />
```


## <a name="contained-in:"></a>Contenu dans :


||
|:-----|
|[CitationText](../../reference/manifest/citationtext.md)|
|[Description](../../reference/manifest/description.md)|
|[DictionaryName](../../reference/manifest/dictionaryname.md)|
|[DictionaryHomePage](../../reference/manifest/dictionaryhomepage.md)|
|[DisplayName](../../reference/manifest/displayname.md)|
|[HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md)|
|[IconUrl](../../reference/manifest/iconurl.md)|
|[QueryUri](../../reference/manifest/queryuri.md)|
|[SourceLocation](../../reference/manifest/sourcelocation.md)|
|[SupportUrl](../../reference/manifest/supporturl.md)|

## <a name="attributes"></a>Attributs



|**Attribut**|**Type**|**Obligatoire**|**Description**|
|:-----|:-----|:-----|:-----|
|Paramètres régionaux|string|obligatoire|Spécifie le nom de culture des paramètres régionaux pour ce remplacement au format de balise de langue BCP 47, comme `"en-US"`.|
|Valeur|string|obligatoire|Spécifie la valeur du paramètre exprimée pour les paramètres régionaux spécifiés.|

## <a name="additional-resources"></a>Ressources supplémentaires



- [Localisation des compléments Office](../../docs/develop/localization.md#off15wecon_LocalesManifest)
    
