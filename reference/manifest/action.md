# <a name="action-element"></a>Action, élément
 Indique l’action à réaliser lorsque l’utilisateur sélectionne des contrôles de [bouton](./control.md#button-control) ou de [menu](./control.md#menu-dropdown-button-controls).
 
## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Oui  | Type d’action à effectuer|


## <a name="child-elements"></a>Éléments enfants

|  Élément |  Description  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Spécifie le nom de la fonction à exécuter. |
|  [SourceLocation](#sourcelocation) |    Spécifie l’emplacement du fichier source pour cette action. |
|  [TaskpaneId](#taskpaneid) | Spécifie l’ID du conteneur de volet des tâches.|
  

## <a name="xsi:type"></a>xsi:type
Cet attribut indique le type d’action réalisée lorsque l’utilisateur sélectionne le bouton. Il peut s’agir de l’une des actions suivantes :
- ExecuteFunction
- ShowTaskpane

## <a name="functionname"></a>FunctionName

Élément obligatoire lorsque **xsi:type** est « ExecuteFunction ». Indique le nom de la fonction à exécuter. La fonction est contenue dans le fichier indiqué dans l’élément [FunctionFile](./functionfile.md).

```xml
<Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation
Élément obligatoire lorsque  **xsi:type** est « ShowTaskpane ». Indique l’emplacement du fichier source pour cette action. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Url** dans l’élément [Urls](./resources.md#urls) dans l’élément [Resources](./resources.md).

```xml
 <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
```  

## <a name="taskpaneid"></a>TaskpaneId
Élément facultatif quand **xsi:type** est « ShowTaskpane ». Spécifie l’ID du conteneur de volet des tâches. Lorsque vous avez plusieurs actions « ShowTaskpane », utilisez un autre **TaskpaneId** si vous souhaitez un volet indépendant pour chacun. Utilisez le même **TaskpaneId** pour différentes actions qui partagent le même volet. Lorsque les utilisateurs choisissent des commandes qui partagent le même **TaskpaneId**, le conteneur de volet reste ouvert, mais le contenu du volet sera remplacé par l’action correspondante « SourceLocation ». 

>**Remarque :** cet élément n’est pas pris en charge dans Outlook.

L’exemple suivant montre deux actions qui partagent le même **TaskpaneId**. 


```xml
 <Action xsi:type="ShowTaskpane">
    <TaskpaneId>MyPane</TaskpaneId>
    <SourceLocation resid="aTaskPaneUrl" />
  </Action>

  <Action xsi:type="ShowTaskpane">
    <TaskpaneId>MyPane</TaskpaneId>
    <SourceLocation resid="anotherTaskPaneUrl" />
  </Action>
```  
