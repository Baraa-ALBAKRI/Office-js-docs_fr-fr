
# <a name="requirements-element"></a>Élément Requirements
Spécifie l’ensemble minimal des conditions requises de l’API JavaScript pour Office ([ensembles des conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md#SpecifyRequirementSets_sets) et/ou méthodes) que votre complément Office doit activer.

 **Type de complément :** Application de contenu, de volet Office, de messagerie


## <a name="syntax:"></a>Syntaxe :


```XML
<Requirements>
   ...
</Requirements>
```


## <a name="contained-in:"></a>Contenu dans :

[OfficeApp](../../reference/manifest/officeapp.md)


## <a name="can-contain:"></a>Peut contenir :



|**Élément**|**Contenu**|**Messagerie**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Ensembles](../../reference/manifest/sets.md)|x|x|x|
|[Méthodes](../../reference/manifest/methods.md)|x||x|

## <a name="remarks"></a>Remarques

Pour plus d’informations sur les ensembles de conditions requises, voir l’article relatif à la [spécification d’hôtes Office et de conditions requises d’API](../../docs/overview/specify-office-hosts-and-api-requirements.md).

