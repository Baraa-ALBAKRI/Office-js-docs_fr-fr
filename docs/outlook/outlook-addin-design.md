# <a name="outlook-add-in-design-guidelines"></a>Instructions de création d’un complément Outlook

Les compléments sont un excellent moyen pour les partenaires d’étendre les fonctionnalités d’Outlook au-delà de notre ensemble de fonctionnalités de base. Les compléments permettent aux utilisateurs d’accéder à des expériences, des tâches et du contenu de tiers sans avoir à quitter leur boîte de réception. Une fois installés, les compléments Outlook sont disponibles sur toutes les plateformes et tous les appareils. Les instructions de haut niveau suivantes vous aideront à concevoir et à créer un complément attrayant, qui apportera le meilleur de votre application directement dans Outlook, sur Windows, le web, iOS, Mac et Android (bientôt disponible).

## <a name="principles"></a>Principes

1. **Concentrez-vous sur quelques tâches clés et exécutez-les correctement**

    Les compléments les mieux conçus sont simples à utiliser, visent un objectif précis et sont réellement utiles pour les utilisateurs. Votre complément s’exécutera dans Outlook, ce principe est donc d’autant plus important. Outlook est une application de productivité : c’est l’endroit où les utilisateurs se rendent pour s’acquitter de leurs tâches.

    Vous allez apporter une extension à notre expérience et vous devez être certain que les scénarios que vous activez s’intègre naturellement au sein d’Outlook. Réfléchissez bien aux situations dans lesquelles la présence des compléments sera le plus utile pour les utilisateurs dans les expériences de messagerie et de calendrier.

    Un complément ne doit pas tenter d’exécuter tout ce que votre application fait déjà. Concentrez-vous sur les actions appropriées les plus fréquemment utilisées, dans le contexte de contenu Outlook. Pensez à votre appel à l’action et indiquez clairement à l’utilisateur ce qu’il doit faire lorsque votre volet de tâches s’ouvre.

2. **Faites en sorte que tout semble aussi naturel que possible**

    Votre complément doit être conçu à l’aide de schémas natifs de la plateforme sur laquelle Outlook s’exécute. Pour ce faire, veillez à respecter et implémenter les instructions d’interaction et visuelles définies par chaque plateforme. Outlook possède ses propres instructions et celles-ci doivent également être prises en compte. Un complément bien conçu sera une combinaison appropriée de votre expérience, de la plateforme et d’Outlook.

    Cela signifie que votre complément devra être visuellement différent lorsqu’il sera exécuté sur Outlook pour iOS et Outlook pour Android (lorsque la prise en charge adéquate est déployée). Nous vous recommandons d’envisager [Framework7](https://framework7.io/) comme une option pour vous aider dans l’application d’un style. Nous publierons des directives mises à jour, en particulier pour Android, dans la mesure où nous nous approchons du lancement de la prise en charge des compléments pour Outlook pour Android.

3. **Faites en sorte que votre complément soit agréable à utiliser jusque dans les moindres détails**

    Nous aimons tous utiliser des produits qui sont à la fois attrayants visuellement et fonctionnels. Pour garantir le succès de votre complément, créez une expérience où chaque interaction et chaque détail visuel ont été soigneusement pensés. Les étapes nécessaires pour exécuter une tâche doivent être claires et pertinentes. Dans l’idéal, aucune action ne doit nécessiter plus d’un ou deux clic(s). Un utilisateur ne doit pas sortir du contexte pertinent pour effectuer une action. Un utilisateur doit pouvoir facilement accéder à votre complément et en sortir pour revenir à ce qu’il faisait avant. Un complément n’est pas destiné à être un emplacement où l’utilisateur passe beaucoup de temps ; il doit s’agir d’une amélioration de nos fonctionnalités principales. Si votre complément est développé correctement, il nous aidera à augmenter la productivité des utilisateurs, ce qui constitue un de nos objectifs.

4. **Personnalisez votre complément à l’image de votre marque de manière judicieuse**

    Nous apprécions les personnalisations et nous savons qu’il est important pour vous de procurer votre expérience unique aux utilisateurs. Cependant, nous pensons que la meilleure façon de garantir la réussite de votre complément est de créer une expérience intuitive qui incorpore subtilement les éléments de votre marque au lieu d’afficher des éléments de marque permanents ou obstruants qui empêchent les utilisateurs de naviguer dans votre système de manière fluide. Vous pouvez par exemple intégrer votre marque en utilisant les couleurs, les icônes et le ton qui la définissent, tout en respectant les modèles privilégiés de la plateforme et les critères d’accessibilité. Efforcez-vous de toujours privilégier le contenu et la capacité à effectuer des tâches plutôt que de chercher à attirer l’attention sur votre marque.

## <a name="design-patterns"></a>Modèles de conception

> **Remarque :** Tandis que les principes ci-dessus s’appliquent à l’ensemble des points de terminaison/plateformes, les modèles et les exemples suivants sont spécifiques des compléments mobiles sur la plateforme iOS.

Pour vous aider à créer un complément bien conçu, nous proposons des [modèles](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/tree/master/Helpful%20Templates/Outlook%20Mobile) pour les versions mobiles avec iOS fonctionnant dans l’environnement Outlook Mobile. Si vous utilisez ces modèles spécifiques, votre complément semblera natif de la plateforme iOS et d’Outlook Mobile. Ces modèles sont également décrits en détail ci-dessous. Bien que cette bibliothèque ne soit pas exhaustive, il s’agit du début de son développement et nous continuerons à l’enrichir à mesure que nous découvrirons des paradigmes que nos partenaires souhaitent inclure dans leurs compléments.  

### <a name="overview"></a>Vue d’ensemble

Un complément type est constitué des éléments suivants.

![Diagramme de modèles d’expérience utilisateur de base pour un volet de tâches sur iOS](../../images/outlook-mobile-design-overview.png)

### <a name="loading"></a>Chargement

Lorsqu’un utilisateur sélectionne votre complément, l’expérience utilisateur doit s’afficher rapidement. Si le chargement est long, utilisez une barre de progression ou un indicateur d’activité. Une barre de progression doit être utilisée lorsque le délai peut être déterminé et un indicateur d’activité doit être utilisé lorsque le délai ne peut pas être déterminé.

![Exemples illustrant une barre de progression et un indicateur d’activité sur iOS](../../images/outlook-mobile-design-loading.png)

### <a name="sign-insign-up"></a>Connexion/Inscription

Votre procédure de connexion (et d’inscription) doit être directe et simple.

![Exemples de pages de connexion et d’inscription sur iOS](../../images/outlook-mobile-design-signin.png)

### <a name="brand-bar"></a>Barre de marque

Le premier écran de votre complément doit inclure un élément de votre marque. Conçue pour que votre marque soit reconnue, la barre de marque vous aide également à définir le contexte pour l’utilisateur. Étant donné que la barre de navigation contient le nom de votre société/marque, il est inutile de reproduire la barre de marque sur les pages suivantes.

![Exemples de barres de marque sur iOS](../../images/outlook-mobile-design-branding.png)

### <a name="margins"></a>Marges

Les marges pour les versions mobiles doivent être définies sur 15 px (8 % de l’écran) de chaque côté, pour s’aligner sur Outlook iOS.

![Exemples de marges sur iOS](../../images/outlook-mobile-design-margins.png)

### <a name="typography"></a>Typographie

La typographie est alignée sur Outlook iOS et doit être simple pour la lisibilité.

![Exemples de typographie pour iOS](../../images/outlook-mobile-design-typography.png)

### <a name="color-palette"></a>Palette de couleurs

L’utilisation des couleurs est subtile dans Outlook iOS.  À des fins de cohérence, nous vous demandons d’utiliser les couleurs uniquement sur les actions et les erreurs, et que seule la barre de marque utilise une couleur unique.

![Palette de couleurs pour iOS](../../images/outlook-mobile-design-color-palette.png)

### <a name="cells"></a>Cellules

Étant donné que la barre de navigation ne peut pas être utilisée pour libeller une page, utilisez les titres de section pour libeller les pages.

![Types de cellules pour iOS](../../images/outlook-mobile-design-cell-types.png)
* * *
![Cellules « Do » pour iOS](../../images/outlook-mobile-design-cell-dos.png)
* * *
![Cellules « Don’t » pour iOS](../../images/outlook-mobile-design-cell-donts.png)
* * *
![Cellules et entrées pour iOS](../../images/outlook-mobile-design-cell-input.png)

### <a name="actions"></a>Actions

Même si votre application gère une multitude d’actions, réfléchissez aux plus importantes que vous souhaitez intégrer à votre complément, et concentrez-vous sur celles-ci.

![Actions et cellules dans iOS](../../images/outlook-mobile-design-action-cells.png)
* * *
![Actions « Do » pour iOS](../../images/outlook-mobile-design-action-dos.png)

### <a name="buttons"></a>Boutons

Les boutons sont utilisés lorsqu’il existe d’autres éléments de l’expérience utilisateur en dessous (par opposition aux actions, car une action est toujours le dernier élément de l’écran).

![Exemples de boutons pour iOS](../../images/outlook-mobile-design-buttons.png)

### <a name="tabs"></a>Onglets

Les onglets peuvent contribuer à organiser le contenu.

![Exemples d’onglets pour iOS](../../images/outlook-mobile-design-tabs.png)

### <a name="icons"></a>Icônes

Les icônes doivent respecter la conception Outlook iOS actuelle autant que possible. Utilisez la taille et la couleur standard.

![Exemples d’icônes pour iOS](../../images/outlook-mobile-design-icons.png)

## <a name="end-to-end-examples"></a>Exemples de bout en bout

Pour le lancement de nos compléments Outlook Mobile v1, nous avons travaillé en étroite collaboration avec nos partenaires qui créaient des compléments. Pour présenter le potentiel de leurs compléments sur Outlook Mobile, notre concepteur a regroupé des flux de bout en bout pour chaque complément, en respectant nos instructions et en utilisant nos modèles.

> **Remarque importante :** ces exemples sont destinés à mettre en évidence la façon idéale de combiner interaction et conception visuelle pour un complément et peuvent ne pas correspondre aux ensembles de fonctionnalités exacts des compléments réels. 

### <a name="giphy"></a>GIPHY

![Conception de bout en bout pour le complément GIPHY](../../images/outlook-mobile-design-giphy.png)

### <a name="nimble"></a>Nimble

![Conception de bout en bout pour le complément Nimble](../../images/outlook-mobile-design-nimble.png)

### <a name="trello"></a>Trello

![Conception de bout en bout pour le complément Trello partie 1](../../images/outlook-mobile-design-trello-1.png)
* * *
![Conception de bout en bout pour le complément Trello partie 2](../../images/outlook-mobile-design-trello-2.png)
* * *
![Conception de bout en bout pour le complément Trello partie 3](../../images/outlook-mobile-design-trello-3.png)

### <a name="dynamics-crm"></a>Dynamics CRM

![Conception de bout en bout pour le complément Dynamics CRM](../../images/outlook-mobile-design-crm.png)
