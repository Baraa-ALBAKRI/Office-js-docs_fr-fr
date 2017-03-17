# <a name="breaktype-javascript-api-for-word"></a>BreakType (JavaScript API for Word)

Spécifie la forme d’un saut.

_S’applique à : Word 2016, Word pour iPad, Word pour Mac, Word Online_

Voici les types de sauts pris en charge par l’API.

| **Valeur**         | **Type** | **Description**     |
|:-----------------|:--------|:----|
|line| | Saut de ligne. |
|page| | Saut de page au point d’insertion.|
|sectionNext| | Saut de section sur la page suivante. Le type de l’élément suivant sera obsolète.|
|sectionContinuous| | Nouvelle section, sans saut de page correspondant.|
|sectionEven| string | Saut de section, la section suivante commençant sur la prochaine page paire. Si le saut de section se produit sur une page paire, Word laisse la prochaine page impaire vide.|
|sectionOdd| string | Saut de section, la section suivante commençant sur la prochaine page impaire. Si le saut de section se produit sur une page paire, Word laisse la prochaine page impaire vide.|

## <a name="support-details"></a>Informations de prise en charge
Utilisez l’[ensemble de conditions requises](../office-add-in-requirement-sets.md) dans les vérifications à l’exécution pour vous assurer que votre application est prise en charge par la version d’hôte de Word. Pour plus d’informations sur la configuration requise pour le serveur et l’application d’hôte Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).
