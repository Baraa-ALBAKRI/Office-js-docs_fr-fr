# <a name="understanding-office-versions"></a>Présentation des versions d’Office

| | |
|:--|:--|
|[![Image de la couverture de l’ouvrage intitulé « Building Office Add-ins using Office.js » (Création de compléments Office à l’aide d’Office.js)](../../images/book-cover.png)](https://leanpub.com/buildingofficeaddins)|**Cet article est un extrait de l’ouvrage « [Building Office Add-ins using Office.js](https://leanpub.com/buildingofficeaddins) » (Création de compléments Office à l’aide d’Office.js) de Michael Zlatkovsky, disponible à l’achat au format électronique sur [LeanPub.com](https://leanpub.com/buildingofficeaddins). (en anglais)**<br/><br/>Copyright © 2016-2017 par Michael Zlatkovsky, tous droits réservés.|

Pour développer et distribuer des compléments qui utilisent le nouveau modèle d’API Office 2016, vous devez avoir Office 2016 ou Office 365 (le sur-ensemble basé sur abonnement qui inclut toutes les fonctionnalités d’Office 2016). Cela semble relativement simple, mais les détails sont un peu plus complexes.


**Le chemin d’accès doré**

Le cas le plus simple possible est lorsque vous (le développeur) et vos utilisateurs finals disposez des versions les plus récentes d’Office 365. Le fait d’avoir Office 365 avec les dernières mises à jour sera certainement le plus simple à des fins de développement et de prototypage.  Si vous êtes un ISV (éditeur de logiciels indépendant) et que vous ne bénéficiez donc d’aucun contrôle sur la version que vos clients utilisent ; ou si vous travaillez au sein d’une entreprise qui n’est pas à la pointe de la technologie, c’est là qu’il devient important de comprendre les versions d’Office.

**Pourquoi c’est important**

Différentes versions (et catégories de versions) proposent différentes zones de surface d’API.  Par exemple, Office 2016 RTM proposait uniquement le premier lot de la nouvelle vague des API Excel et Word. Ces API ont été considérablement développées depuis.  De même, d’autres fonctionnalités, plus particulièrement les commandes de compléments (extensibilité du ruban) et la possibilité de lancer des boîtes de dialogue, n’étaient pas présentes dans la version RTM d’origine.

Dans les pages suivantes, je vais décrire les différentes possibilités d’installation.  Cela peut vous aider à garder l’image suivante à l’esprit :

![Une image qui indique la version MSI d’Office 2016 et l’abonnement Office 365. L’abonnement inclut deux versions : Grand public et Entreprise. La version Grand public a des versions de canal Actuel, Insider slow et Insider fast. La version Entreprise a des versions Canal différé, Première publication du canal différé, canal Actuel et Première publication du canal actuel.](../../images/office-versions.png)


## <a name="office-2016-vs-office-365"></a>Office 2016 et Office 365

Le premier endroit où la zone de surface de l’API bifurque est à la division entre l’installation basée sur MSI d’Office 2016 et l’installation (parfois appelée « Démarrer en un clic ») basée sur un abonnement d’Office 365.

Faisons un point rapide sur Office 365, car j’ai remarqué une certaine confusion concernant le terme.  

Office 365 est un service fondé sur les abonnements qui propose les outils les plus à jour de Microsoft. Il existe des plans Office 365 d’utilisation à domicile et personnelle, ainsi que pour les PME, les grandes entreprises, les écoles et les associations. Tous les plans Office 365 d’utilisation à domicile et d’utilisation personnelle incluent Office 2016 avec les applications Office entièrement installées telles que Word, PowerPoint et Excel, ainsi que le stockage en ligne et bien plus encore. Office 365 pour les utilisateurs professionnels fournit des services de messagerie et de réseaux sociaux via Exchange Server, Skype Entreprise Server, Office Online et l’intégration de Yammer, en plus des logiciels Office.

Ainsi, pour ceux qui viennent du monde SharePoint : oui, SharePoint Online fait partie d’un abonnement Office 365, comme le sont les éditeurs intégrés au navigateur Office Online qui l’accompagnent. Mais, ce n’est pas la *seule* partie de l’abonnement.  L’accès aux *mêmes programmes Office de bureau/Mac que vous connaissez et appréciez* fait également partie de ce même abonnement (comme obtenir les versions iOS et Android de ces programmes Word, Excel, PowerPoint, etc.).


Revenons maintenant aux API : si vous avez Office 2016 Office (sans abonnement), vous aurez *uniquement* l’ensemble initial de la nouvelle vague des API Excel et Word (`ExcelApi 1.1` et `WordApi 1.1`).  Autrement dit : vous aurez uniquement accès à l’*ensemble initial des fonctionnalités d’extensibilité*.  Par conséquent, en plus des améliorations manquantes apportées aux API Excel et Word, il vous manquera aussi d’autres fonctionnalités de complément comme la possibilité de personnaliser le ruban ou de lancer des boîtes de dialogue.

Il est également important de noter que l’offre RTM d’origine des API présente des bogues.  RTM doit être considéré plus comme un *point de départ* vers des API spécifiques de l’hôte enrichies que comme une destination.

Encore une fois : Office 2016, d’un point de vue de l’extensibilité / API, est figé dans le temps... figé sur la fonctionnalité présente au moment où il a été livré en septembre 2015.  

Office 365, quant à lui, est synonyme d’« abonnement ».  Cela se traduit par une version récente et stable (où « stable » pour l’entreprise peut être une version datant de quelques mois ; plus de détails à ce sujet ci-dessous).

Si vous souhaitez accéder aux dernières fonctionnalités API, ce qui est essentiel pour vous, en tant que développeur, vous *devez* être sur une installation basée sur un abonnement d’Office, plutôt que sur l’installation MSI d’Office 2016 figée dans le temps.  En outre, pour la plupart des nouvelles fonctionnalités, vous souhaitez probablement que vos clients soient sur une installation basée sur un abonnement également.


## <a name="office-365-flavors-for-the-consumer"></a>Versions d’Office 365 pour le grand public

Les versions Grand public (autres qu’Entreprise) d’Office 365 comprennent **Office 365 Personnel** et **Office 365 Famille** (avec comme seule différence le nombre d’appareils actifs sur l’abonnement : 1 PC ou Mac, 1 tablette et 1 téléphone, contrairement à 5 de chaque).  Il y a aussi **Office 365 Université**, identique à Office 365 Personnel, mais qui permet l’activation sur *deux* appareils plutôt qu’un.  Pour les trois, la différence réside simplement dans le coût du plan et le nombre d’appareils pris en charge. Ils sont semblables en tout point quant à l’API et à la fonctionnalité.

Les versions Grand public d’Office 365 sont mises à jour chaque mois, avec les mises à jour installées sans assistance et automatiquement. Par conséquent, les versions Grand public d’Office 365, à condition que l’ordinateur soit connecté à Internet, auront toujours accès à la fonctionnalité la plus récente.  Le canal par défaut est le canal « Actuel » (autrement dit, le canal qui est publiquement disponible dans le monde entier), mais l’utilisateur (développeur) audacieux peut également choisir d’être sur l’une des pistes *Insider*.  Il existe deux versions des pistes Insider : Insider Fast et Insider Slow, *Fast* étant vraiment la dernière nouveauté et *Slow* venant quelques plus tard, ancrée sur des versions plus stables.  Dans les deux cas, elles vous permettent de prévisualiser les fonctionnalités à venir un ou deux mois avant le grand public.  Cela peut être particulièrement utile pour les développeurs en vue d’essayer les dernières API avant vos clients, ce qui vous permet de fournir de nouvelles fonctionnalités dès qu’elles sont publiquement disponibles sur les ordinateurs de vos clients.  Associé à l’utilisation du CDN bêta pour Office.js, cela peut également vous permettre de fournir des commentaires sur les API en temps réel à l’équipe, avant qu’elles soient mises en production.  Pour devenir un Insider, voir <https://products.office.com/fr-fr/office-insider>.


## <a name="office-365-flavors-for-enterprise"></a>Versions d’Office 365 Entreprise

Pour les utilisateurs de la version Entreprise d’Office 365 il y a aussi un grand nombre d’options (généralement gérées par l’administrateur informatique).  Comme avec les versions Grand public, il existe un canal « Actuel » (version stable la plus récente, mise à jour tous les mois). De même, il existe une « Première publication du canal actuel », qui est essentiellement identique aux versions « Insider » dans la version consommateur.

Toutefois, les entreprises réticentes à prendre des risques peuvent également choisir d’être sur un canal différé, qui est mis à jour une fois tous les quatre mois au lieu d’une fois par mois.  En outre, ces entreprises peuvent aussi rester sur le canal différé pendant quatre ou même huit mois, avant de passer directement à une version plus récente.  Par conséquent, une entreprise sur le canal différé peut toujours être un peu en retard par rapport au développeur en matière de fonctionnalité d’API disponible (mais moins en retard par rapport à un utilisateur de la version RTM d’Office 2016).


## <a name="office-on-other-platforms-mac-ios-online"></a>Office sur d’autres plateformes (Mac, iOS, Online)

Pour les plateformes autres que PC, un intervalle de temps s’écoule aussi avant que différentes fonctionnalités soient activées.  Cela dépend parfois non seulement du *retard* entre une fonctionnalité complètement codée et mise à la disposition des clients (par exemple, la différence entre Insider et Actuel et Différé) mais aussi de l’ordre dans lequel la fonctionnalité est implémentée sur ces plateformes.  Jusqu’à ce jour, les API Excel sont activées sur la plupart des plateformes à peu près au même moment. Pour Word, la version pour bureau d’Office est généralement en avance par rapport à Office Online.  Pour les fonctionnalités différentes des API (boîtes de dialogue, extensibilité du ruban), elles sont généralement d’abord disponibles pour la version de bureau, puis pour Office Online et Mac.  Les différentes vitesses d’implémentation expliquent pourquoi il est important de garder à l’esprit les versions hôte d’Office mais aussi les versions des API et les ensembles de conditions requises.

>**Cet article est un extrait de l’ouvrage « [Building Office Add-ins using Office.js](https://leanpub.com/buildingofficeaddins) » (Création de compléments Office à l’aide d’Office.js) de Michael Zlatkovsky**. Pour en savoir plus, achetez le livre électronique en ligne sur [LeanPub.com](https://leanpub.com/buildingofficeaddins).

