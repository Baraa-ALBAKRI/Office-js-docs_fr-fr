
# <a name="labs.registerdeserializer"></a>Labs.registerDeserializer

 _**S’applique à :** applications pour Office | Compléments Office | Office Mix | PowerPoint_

Désérialise un objet JSON spécifié dans un objet. Seuls les auteurs de composant doivent l’utiliser.

```
function registerDeserializer(type: string, deserialize: (json: Core.ILabObject): any): void
```


## <a name="parameters"></a>Paramètres


|**Nom**|**Description**|
|:-----|:-----|
|json|Instance [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md) à désérialiser.|

## <a name="return-value"></a>Valeur renvoyée

Renvoie une instance [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md).

