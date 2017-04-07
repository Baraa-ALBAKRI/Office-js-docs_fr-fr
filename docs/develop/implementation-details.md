# <a name="implementation-details-for-those-who-want-to-know-how-it-really-works"></a>Détails d’implémentation, pour en savoir plus sur le fonctionnement *réel*

| | |
|:--|:--|
|[![Image de la couverture de l’ouvrage intitulé « Building Office Add-ins using Office.js » (Création de compléments Office à l’aide d’Office.js)](../../images/book-cover.png)](https://leanpub.com/buildingofficeaddins)|**Cet article est un extrait de l’ouvrage « [Building Office Add-ins using Office.js](https://leanpub.com/buildingofficeaddins) » (Création de compléments Office à l’aide d’Office.js) de Michael Zlatkovsky, disponible à l’achat au format électronique sur [LeanPub.com](https://leanpub.com/buildingofficeaddins). (en anglais)**<br/><br/>Copyright © 2016-2017 par Michael Zlatkovsky, tous droits réservés.|

> *Pendant la rédaction de ce livre, j’ai reçu des commentaires de nouveaux lecteurs et il m’a été demandé d’expliquer plus en détail ce qui se passe en arrière-plan de cette synchronisation/objet proxy.  Si vous êtes curieux et voulez en savoir plus sur les **détails d’implémentation** afin de mieux comprendre le comportement externe d’une API, poursuivez votre lecture.  Sinon, passez à la section suivante.*


## <a name="the-request-context-queue"></a>File d’attente Contexte de la demande

Au cœur de la nouvelle vague des API Office 2016 se trouve un contexte de la demande, qui représente l’objet que vous recevez en tant que paramètre pour la fonction de traitement par lots, à l’intérieur d’une méthode `Excel.run`.  Un objet Contexte de la demande est en fait un référentiel central qui regroupe les modifications que vous souhaitez apporter au document.  Je dis « référentiel » car le contexte de la demande peut en effet être comparé à un système de gestion de version, où tout ce que vous envoyez est le *différences* entre l’état local et l’état distant.

>**Remarque :** Git est une analogie de gestion de version particulièrement bien adaptée, car les modifications locales sont parfaitement isolées du référentiel : tant que vous n’effectuez pas un `git push` de votre état local, le référentiel n’a *aucune connaissance* des modifications apportées.  Les objets proxy et Contexte de la demande du nouveau modèle Office.js sont très similaires : ils sont complètement inconnus du document jusqu'à ce que le développeur émette une commande `context.sync()`.  


L’objet Contexte de la demande contient deux tableaux qui lui permettent de fonctionner.  Un tableau est dédié aux **chemins d’accès d’objet** : descriptions de la dérivation d’un objet à partir d’un autre objet (par exemple, « *appeler la méthode `getRow` avec la valeur du paramètre `2` sur <insert-some-preceding-object-path> pour dériver cet objet* »).  L’autre tableau est dédié aux **actions** (par exemple, *définir la propriété nommée « couleur » sur une valeur « violet » sur l’objet décrit par le chemin d’accès d’objet #xyz*).  Pour ceux qui connaissent bien le modèle de conception « Commande », cette notion de transport d’objets qui représentent la recette d’une action donnée n’est pas nouvelle.

Le Contexte de la demande contient un seul objet racine qui le connecte au modèle objet sous-jacent.  Pour Excel, cet objet est un `workbook` ; pour Word, c’est un `document`.  À partir de là, vous pouvez dériver de nouveaux objets en appelant des méthodes sur l’objet proxy racine ou sur n’importe lequel de ses descendants.  Par exemple, pour obtenir une feuille de calcul nommée « Rapport », demandez l’objet `workbook` pour sa propriété `worksheets` (qui renvoie un objet proxy correspondant à la collection de feuilles de calcul dans le document), puis utilisez `worksheets` pour appeler une méthode `getItem("Report")` afin d’obtenir un objet proxy correspondant à la feuille de calcul « Rapport » souhaitée.  Chacun de ces objets comporte un lien vers son Contexte de la demande d’origine, qui à son tour effectue le suivi des informations sur le chemin de chaque objet : qui était le parent de ce nouvel objet et quelles étaient les circonstances dans lesquelles il a été créé (*était-ce une propriété ou un appel de méthode ? des paramètres ont-ils été transmis ?*).

Chaque fois qu’une méthode ou une propriété est appelée sur un objet proxy, l’appel est enregistré en tant qu’**action** sur le Contexte de la demande. Par exemple, un appel de `range.merge()` ou du paramètre de `fill.color = "purple"` est placé dans la file d’attente en tant qu’action X sur l’objet Y.  En outre, si le résultat de l’appel de la méthode ou de la propriété est un autre objet proxy (par exemple, `worksheets.getItem("Report")` ou `worksheets.add()`), un nouvel objet proxy est généré comme *effet secondaire* de l’appel de la méthode et son lignage sera scrupuleusement indiqué par le Contexte de la demande omniscient.

Examinons un exemple concret.  Supposons que vous avez le code suivant :

**Suivi des opérations objet proxy/Contexte de la demande**
~~~
    Excel.run(async (context) => {
        let range = context.workbook.getSelectedRange();
        range.clear();
        let thirdRow = range.getRow(2);
        firstRow.format.fill.color = "purple";

        await context.sync();
    }).catch(OfficeHelpers.Utilities.log);
~~~

Analysons-le en détail. Avec chaque appel d’API, je noterai le chemin d’accès et leurs actions (exprimés sous une forme abrégée mais suivant étroitement ce qui se produit en interne).

Pour commencer, la ligne **n° 1**  -- `Excel.run(async (context) => {` utilise `Excel.run` pour créer un objet Contexte de la demande. L’appel `.run` effectue un certain nombre d’autres opérations également, mais nous ne nous en occuperons pas pour l’instant. Il est important car il nous donne un tout nouvel objet `context` sur lequel il existe déjà un objet `workbook` pré-initialisé (que nous allons utiliser dans un instant).  

~~~
    objectPaths:
        // markua-start-insert
        1 => global object (workbook)
        // markua-end-insert

    actions: <none>,
~~~


Sur la ligne **n° 2**  -- `let range = context.workbook.getSelectedRange()` : nous utilisons cet objet `workbook` pour dériver un nouvel objet correspondant à la sélection actuelle. Nous l’affectons à une variable appelée `range`, mais il n’est pas important pour le Contexte de la demande : même si nous ne lui avions pas donné de nom et l’avions utilisé en tant qu’objet de transfert pour accéder à une autre destination, il apparaîtrait tout de même dans la liste du Contexte de la demande. La création de l’objet est également indiquée comme action d’initialisation d’objet dans la liste des actions, à des fins décrites plus loin dans cette section.

~~~
    objectPaths:
        P1 => global object (workbook)
        // markua-start-insert
        P2 => (range)
                parent: "P1", type: "method",
                name: "getSelectedRange", args: <none>
        // markua-end-insert

    actions:
        // markua-start-insert
        A1 => action: "init", object: "P2" (range)
        // markua-end-insert
~~~


Ligne **n° 3**  -- `range.clear()` : ajoute la première action réelle ayant un impact sur le document : une commande pour effacer le contenu de la plage :

~~~
    objectPaths:
        P1 => global object (workbook)
        P2 => (range)
                parent: "P1", type: "method",
                name: "getSelectedRange", args: <none>

    actions:
        A1 => action: "init", object: "P2" (range)
        // markua-start-insert
        A2 => action: "method", object: "P2" (range)
                name: "clear", args: <none>
        // markua-end-insert
~~~


Ligne **n° 4**  -- `let thirdRow = range.getRow(2)` : suit un schéma semblable à la ligne n° 2, créant un objet `thirdRow` dérivé de l’objet `range` défini précédemment et ajoutant une autre action d’instanciation :

~~~
    objectPaths:
        P1 => global object (workbook)
        P2 => (range)
                parent: "P1", type: "method",
                name: "getSelectedRange", args: <none>
        // markua-start-insert
        P3 => (thirdRow)
                parent: "P2", type: "method",
                name: "getRow", args: [2]
        // markua-end-insert

    actions:
        A1 => action: "init", object: "P2" (range)
        A2 => action: "method", object: "P2" (range)
                name: "clear", args: <none>
        // markua-start-insert
        A3 => action: "init", object: "P3" (thirdRow)
        // markua-end-insert
~~~


Ligne **n° 5**  -- `firstRow.format.fill.color = "purple"` : est incluse dans plusieurs appels d’API.  Nous commençons par créer un objet format [anonyme], en suivant la propriété `format` de la variable `thirdRow`.  Ensuite, nous procédons de la même façon pour l’objet [anonyme] fill.  Les deux objets suivent le même modèle que précédemment, créant un chemin d’accès d’objet et une action d’instanciation pour chacun.  Mais ensuite, après avoir atteint l’objet souhaité, nous effectuons une autre action ayant un impact sur le document sur l’objet : définir la couleur de remplissage de la troisième ligne en violet (reportez-vous à l’action « **A6** » ci-dessous) :

~~~
    objectPaths:
        P1 => global object (workbook)
        P2 => (range)
                parent: "P1", type: "method",
                name: "getSelectedRange", args: <none>
        P3 => (thirdRow)
                parent: "P2", type: "method",
                name: "getRow", args: [2]
        // markua-start-insert
        P4 => (format)
                parent: "P3", type: "property",
                name: "format"
        P5 => (fill)
                parent: "P4", type: "property",
                name: "fill"
        // markua-end-insert

    actions:
        A1 => action: "init", object: "P2" (range)
        A2 => action: "method", object: "P2" (range)
                name: "clear", args: <none>
        A3 => action: "init", object: "P3" (thirdRow)
        // markua-start-insert
        A4 => action: "init", object: "P4" (format)
        A5 => action: "init", object: "P5" (fill)
        A6 => action: "setter", object: "P5" (fill),
                name: "color", value: "purple"
        // markua-end-insert
~~~


Pour finir, sur la ligne **n° 7** (la ligne n° 6 était vide), nous accédons à l’incantation **`await context.sync()`** magique.  Cette commande indique à l’objet Contexte de la demande de regrouper toutes les informations pertinentes (actions en attente et toutes les informations relatives au chemin d’accès d’objet associées) et de les envoyer à l’application hôte pour le traitement.


Aux extrémités de réception de l’application hôte, l’hôte individualise les actions et commence à les parcourir une par une.  Il conserve un dictionnaire des objets qui ont été dérivés lors de cette session `sync` particulière. Ainsi, après avoir récupéré une fois la plage correspondant à `thirdRow`, il sera inutile de la réévaluer.  Ceci améliore l’efficacité et empêche les erreurs : cela vous évite de récupérer à nouveau la ligne à l’index 2 correspondant si d’autres lignes ont été ajoutées entre ce dernier et la première ligne. Cela évite aussi de récupérer à nouveau la sélection chaque fois, car elle a pu être déplacée (par exemple, lors de l’ajout et l’activation d’une feuille de calcul), mais sémantiquement, la plage doit être *imprimée* avec la référence d’origine.  Enfin, si vous disposez d’un objet dérivé de l’appel à la méthode `add` sur l’objet worksheet-collection, vous ne souhaitez *absolument* pas dériver de nouveau l’objet (c’est-à-dire, ajouter une nouvelle feuille) à chaque fois que vous accédez à l’objet.

Si un problème survient au cours du processus, le reste du lot est annulé. En continuant avec l’exemple précédent, s’il n’existe aucune troisième ligne dans la sélection (autrement dit, il s’agit d’une sélection de cellules 2x2), les commandes restantes sont ignorées (ce à quoi vous vous attendiez probablement). La chose importante, cependant, est qu’il n’y aucune *atomicité* à l’action `Excel.run` ou `sync` : toutes les actions qui ont déjà été effectuées sont *définitivement* terminées.  Dans cet exemple, le document peut être laissé dans un état *modifié*, où l’effacement de la sélection a déjà eu lieu, mais la mise en forme de la troisième ligne n'a pas encore été effectuée.  Ce n’est pas l’idéal mais cela ne diffère pas de VBA ou VSTO en ce qui concerne l’automatisation Office ; la restauration est tout simplement trop difficile, notamment en raison des actions utilisateur ou collaborateur qui ont pu avoir lieu entre-temps. 

## <a name="the-host-applications-response"></a>Réponse de l’application hôte

Supposons que la `sync` a réussi : chaque objet nécessaire (la sélection d’origine, sa troisième ligne, la mise en forme, le remplissage) a été créé correctement et les deux actions ayant un impact sur le document ont également pu être validées dans le document.  Et après ?

Comme mentionné précédemment, l’hôte conserve un dictionnaire de l’objet qu’il a utilisé.  Toutefois, ce dictionnaire est valable *uniquement pour la durée de la `sync` en question : pas pour la durée de vie de l’application*.  Réussir à conserver et suivre les objets indéfiniment serait un immense succès.

Prenons maintenant le cas où un chemin d’accès d’objet est l’action « Ajouter » sur une collection de feuilles de calcul.   Lors du traitement de la `sync`, la méthode aurait été exécutée une seule fois (avec l’effet secondaire approprié de la création de la feuille de calcul) et la feuille résultante aurait été mise en cache.  C’est formidable pour la `sync` actuelle, mais que se passe-t-il si le développeur souhaite de nouveau accéder à la feuille lors d’une prochaine `sync` ?  C’est ici que les actions d’instanciation mentionnées précédemment entrent en jeu.

Pour chaque action, l’application hôte peut *éventuellement* envoyer une réponse.  Pour les actions telles que l’effacement d’une plage ou la définition d’une couleur de remplissage, il n’y a pas de réponse à donner (le fait que l’opération a réussi est évident puisque l’exécution de la file d’attente s’est poursuivie jusqu'à la fin).  Mais pour les actions d’instanciation, l’hôte *peut* envoyer une réponse pour indiquer à JavaScript de remapper son chemin d’accès d’objet à un élément moins volatile. Par conséquent, tandis que le chemin d’accès d’origine d’une feuille nouvellement créée a peut-être été « *exécuter la méthode `add` sur l’objet xyz* » (où xyz est la collection de feuilles de calcul), la réponse peut indiquer « *à partir d’ici, se reporter à la feuille comme étant un appel « getItem » avec le paramètre « 123456789 » sur le même objet xyz ».  Autrement dit, en créant l’objet et en exécutant l’action instanciation, l’hôte peut déterminer s’il existe un ID plus permanent qu’il peut renvoyer à JavaScript pour des  références ultérieures à cet objet.  (Un exemple moins radical : l’extraction d’une feuille par son nom est un peu risquée, dans la mesure où les noms peuvent changer, via l’interaction de l’utilisateur et par programme. Cependant, si l’hôte peut remapper le chemin d’accès à un ID de feuille de calcul permanent, les appels ultérieurs sur l’objet continueront à faire référence à la même feuille, indépendamment de son nom).

Il existe toutefois une autre utilisation encore plus importante pour les réponses de l’hôte.  Supposons que, du côté JavaScript, vous avez un appel à `range.load("formulas")`. En matière d’actions, cela est représenté par une action *query* sur l’objet, avec un paramètre dont la valeur est « formules ». L’hôte répondra à cette action par l’extraction de l’objet approprié (qui se trouve déjà dans son dictionnaire, grâce à l’action d’instanciation), en l’interrogeant pour les propriétés requises et en renvoyant les informations demandées.


## <a name="back-on-the-proxy-objects-territory"></a>Retour sur le territoire de l’objet proxy

De retour dans JavaScript, la `sync` attend patiemment une réponse de l’application hôte.  Le code du développeur attend *aussi* patiemment, en utilisant un `await` ou en s’abonnant à l’appel de fonction `.then` de la promesse `sync`.

Lorsque la réponse est *renvoyée*, un petit traitement interne a lieu avant que l’exécution ne retourne au code du développeur.  Par exemple, les remappages de chemin d’accès éventuels décrits dans la section précédente prennent effet.  Un traitement interne a lieu également (par exemple, les chemins d’accès des objets qui étaient valides au cours du lot `sync` précédent deviennent non valides mais ne peuvent plus être utilisés, je donnerai plus d’explications à ce sujet).  Et surtout, les résultats des actions *query* prennent effet, en utilisant les valeurs chargées et en les plaçant de nouveau sur les propriétés et les objets correspondants.  Cela garantit que, après la `sync`, si le code du développeur fait maintenant référence à `range.values` pour une plage dont les valeurs ont été chargées, il obtiendra le dernier instantané connu des valeurs (au lieu d’une erreur `PropertyNotLoaded`).

Lorsque le post-traitement est terminé, le Contexte de la demande peut être réutilisé.  Son tableau des actions a été rétabli sur une page blanche au tout début de la `sync`. À l’inverse, les chemins d’accès d’objets du tableau des chemins d’accès d’objets (qui n’est jamais vidé pendant la durée de vie de l’objet Contexte de la demande en question, car les actions ultérieures sont liées à la réutilisation de certains des chemins existants) ont été modifiés, en fonction des réponses de l’hôte et du post-traitement.  Par conséquent, un nouveau lot d’opérations peut commencer, mis en attente jusqu'à la prochaine `await context.sync()`.


## <a name="a-special-but-common-case-objects-without-ids"></a>Cas spécial (mais courant) : objets sans ID

Lorsque vous travaillez avec des objets, tels que des feuilles de calcul (Excel) ou des contrôles de contenu (Word), le travail de l’application hôte est relativement simple : dans les deux cas, il existe un ID permanent associé à chacun de ces objets, donc peu importe comment l’objet a été créé (`getActiveWorksheet()`, ou `getItem`, ou tout autre appel), l’hôte peut toujours utiliser l’action d’instanciation pour remapper le chemin d’accès à un ID permanent.  Ce qui signifie que, en tant que développeur, après avoir créé l’objet une fois à un moment donné, vous pouvez continuer à l’utiliser dans la prochaine `sync`, ou encore plus longtemps par la suite.  Pas de surprise.

Mais qu’en est-il des objets qui n’ont pas d’ID ; et qui, par définition, possèdent un nombre infini de permutations à leur sujet ?  Il n’est pas du tout facile d’obtenir une référence concrète aux plages Excel (un regroupement spécifique de cellules) et aux plages Word (du texte commençant à un endroit et finissant à un autre) : l’adresse/l’index auxquels elles se trouvent peut être déplacé, et les plages peuvent aussi s’agrandir et se développer si des cellules ou des caractères supplémentaires y sont ajoutés.  Il en est de même pour d’autres objets.

L’hôte ne rencontre pas de problèmes de suivi des objets lors du *traitement* du lot, car il peut utiliser un dictionnaire interne et (après avoir récupéré l’objet une fois) continuer à l’utiliser pendant la durée du lot.  Toutefois, comme indiqué précédemment, l’hôte est *sans état* dans les lots : il ne peut pas conserver une référence à chaque objet qui a fait l’objet d’un accès depuis JavaScript. En effet, l’application subira une fuite de mémoire et s’arrêtera.

Pour éviter de paralyser l’application hôte avec des milliers d’objets inutiles tout en permettant d’effectuer un `load` et une `sync` (scénario très fréquent) avant d’effectuer d’autres actions avec un objet, la conception d’origine était comme suit :

1. Par défaut, tout objet sans ID perd la connexion avec le document sous-jacent sans assistance et ne peut pas être réutilisé après la `sync` actuelle. Il peut toujours être lu par JavaScript *après* la synchronisation (autrement il serait absurde de le charger en premier lieu) et cela peut convenir pour certains scénarios, mais ce comportement par défaut exclurait, par exemple, l’exemple de mise en surbrillance dans lequel les valeurs de la plage sont d’abord lues puis colorées en conséquence.
2. Pour permettre le dernier scénario, nous exposons une API (`context.trackedObjects.add`) où un développeur peut explicitement dire à l’hôte « *Je veux effectuer le suivi de cet objet à plus long terme, pas seulement pendant la `sync`* actuelle.  Le développeur est alors responsable d’appeler `context.trackedObjects.remove` lorsque l’objet n’est plus nécessaire (aucune pénalité sévère s’il ne le fait pas mais cela ralentira l’application hôte au fil du temps, utilisez donc le suivi des objets avec modération et arrêtez de l’utiliser dès que vous avez terminé).

Sur le calque JavaScript, un appel à `context.trackedObjects.add` ajouterait un nouveau type d’action à la file d’attente, indiquant que l’objet avec l’ID X voudrait être suivi.  Du côté de l’hôte, cette action serait interprétée pour créer une classe wrapper permanente autour de l’objet en mémoire, créant un ID que l’objet pourrait utiliser comme s’il s’agissait d’un ID réel. Cet ID serait renvoyé vers l’objet, tout comme le résultat du remappage d’un chemin d’accès d’objet d’une action d’instanciation.  Et de la même manière, un appel à `context.trackedObjects.remove` obtiendrait aussi une action spéciale ajoutée à la file d’attente, demandant que l’hôte libère la mémoire pour l’objet devenu inutile et marquant l’objet lui-même comme n’ayant plus un chemin d’accès valide.

Cette conception a fonctionné (et fonctionne encore aujourd’hui) si un développeur décide de créer un objet Contexte de la demande manuellement, via `var context = new Excel.RequestContext()`, au lieu de `.run`.  Dans la pratique, à la fois dans notre test interne et dans la version publique, il s’est avéré très fastidieux d’avoir à appeler `context.trackedObjects.add` sur un objet ou deux dans presque chaque scénario.  Et même lorsque les développeurs l’appelaient (avec des essais et erreurs), cela était encore plus fastidieux (non, irréaliste) de s’attendre à ce que les gens se souviennent et se débarrassent correctement des objets suivis devenus inutiles.

En observant les utilisateurs se confronter à ce concept d’objets suivis, il est devenu évident que dans la plupart des cas, l’intention du développeur n’est *pas* de conserver l’objet pour un stockage à long terme mais plutôt de le suivre pour pouvoir l’utiliser sur une ou deux limites `sync`, tout simplement.  C’est là que la méthode `Excel.run` (`Word.run`, etc.) est née : pour permettre aux développeurs de déclarer une seule unité sémantique d’automatisation, même si en interne elle est constituée d’une série de `sync`.  Et également pour que l’infrastructure gère le suivi et l’annulation du suivi sans assistance.

Cela signifie que lorsque vous lancez un `Excel.run` (`Word.run`, etc.), après chaque action d’instanciation, il y a *aussi* une action pour effectuer le suivi de l’objet.  Tout à la fin, après avoir vidé la file d’attente à la fin du `Excel.run`, une demande interne distincte est effectuée pour annuler le suivi de chaque objet non dérivable et sans ID créé entre-temps.  Par conséquent l’image *réelle* du tableau des « actions » indiquée plus haut est en fait un peu plus détaillée :

~~~
    actions:
        A1 => action: "init", object: "P2" (range)
        // markua-start-insert
        A2 => action: "track", object: "P2" (range)
        // markua-end-insert
        A3 => action: "method", object: "P2" (range)
                name: "clear", args: <none>
        A4 => action: "init", object: "P3" (thirdRow)
        // markua-start-insert
        A5 => action: "track", object: "P3" (thirdRow)
        // markua-end-insert
        A6 => action: "init", object: "P4" (format)
        A7 => action: "init", object: "P5" (fill)
        A8 => action: "setter", object: "P5" (fill),
                name: "color", value: "purple"
~~~


Puis, à la fin du `run`, après avoir vidé la file d’attente, les éléments suivants sont envoyés (notez que seuls les chemins d’accès d’objet appropriés sont envoyés ; il est inutile de surcharger la limite de processus) :


~~~
    objectPaths:
        P1 => global object (workbook)
        P2 => (range)
                parent: "P1", type: "method",
                name: "getSelectedRange", args: <none>
        P3 => (thirdRow)
                parent: "P2", type: "method",
                name: "getRow", args: [2]

    actions:
        A1 => action: "untrack", object: "P2" (range)
        A2 => action: "untrack", object: "P3" (thirdRow)
~~~


En bref, c’est ainsi que fonctionnent les objets proxy sous-jacents et que l’exécution gère ses communications vers et depuis l’application hôte.  

>**Cet article est un extrait de l’ouvrage « [Building Office Add-ins using Office.js](https://leanpub.com/buildingofficeaddins) » (Création de compléments Office à l’aide d’Office.js) de Michael Zlatkovsky**. Pour en savoir plus, achetez le livre électronique en ligne sur [LeanPub.com](https://leanpub.com/buildingofficeaddins).