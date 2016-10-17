
# <a name="sets-element"></a>Élément Sets
Spécifie le sous-ensemble minimal de l’API JavaScript pour Office nécessaire à l’activation de votre complément Office.

 **Type de complément :** Application de contenu, de volet Office, de messagerie


## <a name="syntax:"></a>Syntaxe :


```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```


## <a name="contained-in:"></a>Contenu dans :

[Requirements](../../reference/manifest/requirements.md)


## <a name="can-contain:"></a>Peut contenir :

[Ensemble](../../reference/manifest/set.md)


## <a name="attributes"></a>Attributs



|**Attribut**|**Type**|**Obligatoire**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|chaîne|facultatif|Spécifie la valeur de l’attribut **MinVersion** par défaut pour tous les éléments [Set](../../reference/manifest/set.md) enfants. La valeur par défaut est « 1.1 ».|

## <a name="remarks"></a>Remarques

Pour plus d’informations sur les ensembles de conditions requises, voir l’article relatif à la [spécification d’hôtes Office et de conditions requises d’API](../../docs/overview/specify-office-hosts-and-api-requirements.md).

Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et sur l’attribut **DefaultMinVersion** de l’élément **Sets**, voir l’article relatif à la [définition de l’élément Requirements dans le manifeste](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).

