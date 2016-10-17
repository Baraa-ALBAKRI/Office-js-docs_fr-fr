## <a name="supertip"></a>Supertip
Définit une info-bulle enrichie (titre et description). Il est utilisé par les contrôles de [bouton](./button.md) et de [menu](./menu-control.md). 

## <a name="child-elements"></a>Éléments enfants
|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [Title](#title)        | Oui |   Texte de l’info-bulle.         |
|  [Description](#description)  | Oui |  Description de l’info-bulle.    |

## <a name="title"></a>Titre
Obligatoire. Texte de la propriété SuperTip. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément [ShortStrings](./resources.md#shortstrings) dans l’élément [Resources](./resources.md).

## <a name="description"></a>Description
Obligatoire. Description de la propriété SuperTip. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément [LongStrings](./resources.md#longstrings) dans l’élément [Resources](./resources.md).

```xml
 <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
```