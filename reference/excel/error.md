# <a name="officeextension.error-object-(javascript-api-for-excel)"></a>Objet OfficeExtension.Error (API JavaScript pour Excel)

Représente les erreurs qui se produisent lorsque vous utilisez l’API JavaScript Excel.

_S’applique à : Excel 2016, Excel Online, Excel pour iOS, Office 2016_

## <a name="properties"></a>Propriétés
| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|code|chaîne|Obtient une valeur qui indique le type d’erreur. La valeur peut être « AccessDenied », « ActivityLimitReached », « BadPassword », « GeneralException », « InsertDeleteConflict », « InvalidArgument », « InvalidBinding », « InvalidOperation », « InvalidReference », « InvalidSelection », « ItemAlreadyExists », « ItemNotFound », « NotImplemented » ou « UnsupportedOperation ». |
|debugInfo|string|Obtient une valeur qui indique ce qui s’est passé lorsque l’erreur est survenue. Cette valeur est uniquement destinée au développement/débogage.  |
|message |string| Obtient une chaîne localisée explicite qui correspond au code d’erreur.|
|name |string| Obtient une valeur qui est toujours « OfficeExtension.Error ». |
|traceMessages |string[]| Obtient un tableau de valeurs qui correspondent aux messages d’instrumentation définis avec context.trace(); |

## <a name="methods"></a>Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[toString()](#tostring)|chaîne|Renvoie le code d’erreur et le message au format suivant : « {0}: {1} », code, message.|

## <a name="method-details"></a>Détails de méthodes

### <a name="tostring()"></a>toString()
Renvoie le code d’erreur et le message au format suivant : « {0}: {1} », code, message.

#### <a name="syntax"></a>Syntaxe
```js
error.toString()
```

#### <a name="parameters"></a>Paramètres
Aucun

#### <a name="returns"></a>Retourne
string
