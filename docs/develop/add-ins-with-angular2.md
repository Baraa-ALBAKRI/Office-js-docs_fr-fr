# <a name="tips-for-creating-office-add-ins-with-angular-2"></a>Conseils pour la création de compléments Office avec Angular 2 

Cet article fournit des conseils sur l’utilisation d’Angular 2 pour créer un complément Office sous la forme d’une application monopage.

>**Remarque :** Avez-vous une contribution à apporter à partir de votre expérience d’utilisation d’Angular 2 pour créer des compléments Office ? Vous pouvez contribuer à cet article dans [GitHub](https://github.com/OfficeDev/office-js-docs) ou fournir vos commentaires en envoyant un [problème](https://github.com/OfficeDev/office-js-docs/issues) dans le référentiel. 

Pour un exemple de complément Office créé à l’aide de l’infrastructure Angular 2, consultez [Complément de vérification du style dans Word basé sur Angular 2](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## <a name="bootstrapping-must-be-inside-officeinitialize"></a>L’amorçage doit s’effectuer à l’intérieur d’Office.initialize

Dans une page qui appelle les API Office, Word ou Excel JavaScript, votre code doit d’abord attribuer une méthode à la propriété `Office.initialize`. (Si vous ne possédez aucun code d’initialisation, le corps de la méthode peut contenir simplement des symboles « `{}` » vides, mais vous ne devez pas laisser la propriété `Office.initialize` non définie. Pour plus d’informations, voir [Initialisation de votre complément](http://dev.office.com/docs/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in).) Office appelle cette méthode immédiatement après l’initialisation des bibliothèques JavaScript Office.

**Votre code d’amorçage Angular doit être appelé à l’intérieur de la méthode que vous affectez à `Office.initialize`** pour vous assurer que les bibliothèques JavaScript Office ont été initialisées en premier. Voici un exemple simple qui montre comment procéder. Ce code doit figurer dans le fichier main.ts du projet.

```js
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
    import { AppModule } from './app.module';
    Office.initialize = function () {
        const platform = platformBrowserDynamic();
        platform.bootstrapModule(AppModule);
  };
```

## <a name="use-the-hash-location-strategy-in-the-angular-application"></a>Utiliser la stratégie d’emplacement de hachage dans l’application Angular

La navigation entre des itinéraires dans l’application peut ne pas fonctionner si vous ne spécifiez pas la stratégie d’emplacement de hachage. Vous pouvez procéder de deux manières. Tout d’abord, vous pouvez spécifier un fournisseur pour la stratégie d’emplacement dans le module de votre application, comme montré dans l’exemple suivant. Il est placé dans le fichier app.module.ts.

```js
import { LocationStrategy, HashLocationStrategy } from '@angular/common';
// Other imports suppressed for brevity
    @NgModule({
        providers: [
            {provide: LocationStrategy, useClass: HashLocationStrategy},
            // Other providers suppressed
        ],
        // Other module properties suppressed
  })
  export class AppModule {}
``` 

Si vous définissez vos itinéraires dans un module de routage distinct, il existe une autre façon de spécifier la stratégie d’emplacement de hachage. Dans le fichier .ts de votre module de routage, passez un objet de configuration vers la fonction `forRoot` qui spécifie la stratégie. Voici un exemple de code. 

```js
import { RouterModule, Routes } from '@angular/router';
// Other imports suppressed for brevity
    const routes: Routes = // route definitions go here
    @NgModule({
      imports: [ RouterModule.forRoot(routes, {useHash: true}) ],
      exports: [ RouterModule ]
    })
    export class AppRoutingModule {}
```   


## <a name="consider-wrapping-fabric-components-with-angular-2-components"></a>Insertion de composants Fabric dans des composants Angular 2

Nous vous recommandons d’utiliser le style [Office UI Fabric](http://dev.office.com/fabric#/fabric-js) dans votre complément. Fabric comprend des composants disponibles dans plusieurs versions, y compris une version qui [repose sur TypeScript](https://github.com/OfficeDev/office-ui-fabric-js). Envisagez d’utiliser des composants Fabric dans votre complément en les insérant dans des composants Angular 2. Pour obtenir un exemple de la procédure à suivre, consultez l’article relatif au [complément de vérification du style Word reposant sur Angular 2](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker). Observez, par exemple, comment le composant Angular défini dans [fabric.textfield.wrapper](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component.ts) importe le fichier Fabric TextField.ts, dans lequel le composant Fabric est défini. 


## <a name="using-the-office-dialog-api-with-angular"></a>Utilisation de l’API Boîte de dialogue Office

L’API Boîte de dialogue du complément Office permet à votre complément d’ouvrir une page dans une boîte de dialogue semi-modale dans laquelle vous pouvez échanger des informations avec la page principale, qui se trouve généralement dans un volet Office. 

La méthode [displayDialogAsync](http://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) accepte un paramètre qui indique l’URL de la page qui doit s’ouvrir dans la boîte de dialogue. Votre complément peut avoir une autre page HTML (différente de la page de base) pour passer à ce paramètre, ou vous pouvez passer l’URL d’un itinéraire dans votre application Angular. 

Il est important de ne pas oublier, si vous passez un itinéraire, que la boîte de dialogue crée une nouvelle fenêtre avec son propre contexte d’exécution. Votre page de base et son code d’initialisation et d’amorçage s’exécutent à nouveau dans ce nouveau contexte, et toutes les variables sont définies sur leurs valeurs initiales dans la boîte de dialogue. Par conséquent, cette technique lance une deuxième instance de votre application monopage dans la boîte de dialogue. Le code qui modifie des variables dans la boîte de dialogue ne change pas la version du volet Office des mêmes variables. De même, la boîte de dialogue possède son propre stockage de session, qui n’est pas accessible à partir du code dans le volet Office.  


## <a name="forcing-an-update-of-the-dom"></a>Forcer une mise à jour du modèle DOM

Dans n’importe quelle application Angular 2, il arrive que les notifications de mise à jour de DOM ne se déclenchent pas. L’infrastructure fournit une méthode `tick()` sur l’objet `ApplicationRef` qui force une mise à jour. Voici un exemple de code.

```js
import { ApplicationRef } from '@angular/core';
    export class MyComponent {
        constructor(private appRef: ApplicationRef) {}
        myMethod() {
            // Code that changes the DOM is here
            appRef.tick();
        }
}
``` 

## <a name="using-observables"></a>Utilisation d’éléments visibles

Angular 2 utilise RxJS (Reactive Extensions for JavaScript), et RxJS présente les objets `Observable` et `Observer` pour implémenter le traitement asynchrone. Cette section fournit une brève introduction à l’utilisation de `Observables` ; pour plus d’informations, consultez la documentation [RxJS](http://reactivex.io/rxjs/) officielle.

Un `Observable` est semblable à un objet `Promise` d’une certaine façon - il est renvoyé immédiatement à partir d’un appel asynchrone, mais il ne peut être résolu qu’après un certain délai. Toutefois, bien qu’une `Promise` soit une valeur unique (qui peut être un objet de tableau), un `Observable` est un tableau d’objets (éventuellement avec un seul membre). Cela permet d’appeler les [méthodes de tableaux](http://www.w3schools.com/jsref/jsref_obj_array.asp), telles que `concat`, `map` et `filter`, sur des objets `Observable`. 

### <a name="pushing-instead-of-pulling"></a>Poussée au lieu d’extraction

Votre code « pousse » les objets `Promise` en les affectant aux variables, mais les objets `Observable` « poussent » leurs valeurs vers les objets qui *s’abonnent* à l’objet `Observable`. Les abonnés sont des objets `Observer`. L’avantage de l’architecture Push est que les nouveaux membres peuvent être ajoutés au tableau `Observable` au fil du temps. Lorsqu’un nouveau membre est ajouté, tous les objets `Observer` qui s’abonnent à `Observable` reçoivent une notification. 

L’`Observer` est configuré pour traiter chaque nouvel objet (appelé l’objet « suivant ») avec une fonction. (Il est également configuré pour répondre à une erreur et à une notification d’achèvement. Consultez la section suivante pour obtenir un exemple.) Pour cette raison, les objets `Observable` peuvent être utilisés dans un plus large éventail de scénarios que les objets `Promise`. Par exemple, en plus de retourner un `Observable` à partir d’un appel AJAX, de la façon dont vous pouvez retourner une `Promise`, un `Observable` peut être renvoyé à partir d’un gestionnaire d’événements, tel que le gestionnaire d’événements « modifié » pour une zone de texte. Chaque fois qu’un utilisateur saisit du texte dans la zone, tous les objets `Observer` abonnés réagissent immédiatement en utilisant le dernier texte et/ou l’état actuel de l’application en tant qu’entrée. 


### <a name="waiting-until-all-asynchronous-calls-have-completed"></a>Attendre jusqu'à ce que tous les appels asynchrones soient terminés

Lorsque vous voulez vous assurer qu’un rappel ne s’exécute que lorsque tous les membres d’un ensemble d’objets `Promise` sont résolus, utilisez la méthode `Promise.all()`.

```js
myPromise.all([x, y, z]).then(// TODO: Callback logic goes here.)
``` 

Pour faire la même chose avec un objet `Observable`, vous utilisez la méthode [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md).  

```js
var source = Rx.Observable.forkJoin([x, y, z]);

var subscription = source.subscribe(
  function (x) {
    // TODO: Callback logic goes here
  },
  function (err) {
    console.log('Error: ' + err);
  },
  function () {
    console.log('Completed');
  });
``` 

