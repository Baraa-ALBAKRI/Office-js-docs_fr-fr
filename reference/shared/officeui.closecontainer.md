# <a name="uiclosecontainer-method"></a>Méthode UI.closeContainer

Cette méthode ferme le conteneur d’IU où le code JavaScript est exécuté. Le comportement de cette méthode est spécifié dans le tableau suivant.

| Lorsqu’elle est appelée depuis | Comportement |
|:-----------------|:---------|
| Un bouton de commande sans IU | Aucun effet. Les boîtes de dialogue ouvertes par [displayDialogAsync](officeui.displaydialogasync.md) restent ouvertes. |
| Un volet Office | Le volet Office se ferme. Les boîtes de dialogue ouvertes par `displayDialogAsync` se ferment également. Si le volet Office est épinglable et qu’il a été épinglé par l’utilisateur, il sera détaché. |
| Une extension de module | Aucun effet. |

## <a name="syntax"></a>Syntaxe

```js
Office.context.ui.closeContainer();
```

## <a name="returns"></a>Renvoie
void