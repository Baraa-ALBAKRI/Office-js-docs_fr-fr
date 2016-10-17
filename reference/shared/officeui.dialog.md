#<a name="ui.dialog-object"></a>Objet UI.Dialog
Objet renvoyé lorsque la méthode [displayDialogAsync](officeui.displaydialogasync.md) est appelée.

## <a name="members"></a>Membres
| Membre	       | Type	   |Description|
|:---------------|:--------|:----------|
|fermer|Fonction|Permet au complément de fermer sa boîte de dialogue.|
|addEventHandler|Fonction|Enregistre un gestionnaire d’événements. Les deux événements suivants sont pris en charge : <ul><li>DialogMessageReceived. Déclenché lorsque la boîte de dialogue envoie un message à son parent.</li><li>DialogEventReceived. Déclenché lorsque la boîte de dialogue a été fermée ou lorsque son chargement a été annulé.</li></ul> |


### <a name="close()"></a>close()
Appelé à partir d’une page parent pour fermer la boîte de dialogue correspondante.     
```js    
[dialogObject].close();    
``` 

#### <a name="parameters"></a>Paramètres    
Aucun 

#### <a name="returns"></a>Retourne    
void  


#### <a name="examples"></a>Exemples
Pour obtenir des exemples, consultez la rubrique [Méthode DisplayDialogAsync](officeui.displaydialogasync.md).
