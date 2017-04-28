
# <a name="supporturl-element"></a>SupportUrl, élément

Spécifie l’URL d’une page qui fournit des informations de prise en charge pour votre complément.

## <a name="example"></a>Exemple

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>

```


## <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**Obligatoire**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|obligatoire|Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](../../reference/manifest/defaultlocale.md).|

## <a name="child-elements"></a>Éléments enfants

|  Élément | Obligatoire | Description  |
|:-----|:-----|:-----|
|  [Override](../../reference/manifest/override.md)   | Non | Spécifie le paramètre pour les URL de paramètres régionaux supplémentaires |

## <a name="parent-element"></a>Élément parent
[OfficeApp](../../reference/manifest/officeapp.md)

