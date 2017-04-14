# <a name="use-centralized-deployment-to-publish-office-add-ins"></a>Utilisation du déploiement centralisé pour publier des compléments Office

Le centre d’administration Office 365 permet aux administrateurs de déployer facilement des compléments Word, Excel et PowerPoint auprès d’utilisateurs ou de groupes au sein de leur organisation. Les compléments déployés via le centre d’administration sont disponibles pour les utilisateurs directement dans leurs applications Office, sans qu’aucune configuration client ne soit requise. Vous pouvez déployer des compléments internes, ainsi que des compléments fournis par des éditeurs de logiciels indépendants via le déploiement centralisé.

Le centre d’administration prend actuellement en charge les scénarios suivants :

- Déploiement centralisé de nouveaux compléments et de ceux mis à jour pour des utilisateurs, des groupes ou une organisation.
- Déploiement de plusieurs plateformes, y compris Windows et Office Online (Mac bientôt disponible).
- Déploiement en anglais et pour les clients du monde entier.
- Déploiement de compléments hébergés sur le cloud.
- Installation automatique de l’application Office au lancement.
- URL des compléments hébergées au sein d’une zone protégée par un pare-feu.
- Déploiement de compléments Office Store (bientôt disponible).

<!--
The admin center also includes a pre-deployment validation checking service.
-->

Les investissements futurs dans des scénarios de déploiement de compléments porteront sur le centre d’administration Office 365. Nous vous recommandons d’utiliser le centre d’administration pour déployer des compléments dans votre organisation, si votre organisation remplit les conditions préalables.

## <a name="prerequisites-for-centralized-deployment"></a>Conditions préalables au déploiement centralisé 

Vous pouvez déployer des compléments via le centre d’administration si votre organisation répond aux critères suivants :

- Les utilisateurs exécutent une version d’Office 2016 ProPlus :
    - Windows version 16.0.8027 ou ultérieure
    - Mac version 15.33.170327 ou ultérieure
- Les utilisateurs se connectent à Office 2016 avec leur compte professionnel ou scolaire.
- Votre organisation utilise le service d’identité Azure Active Directory (Azure AD).
- La méthode d’authentification [OAuth est activée](https://msdn.microsoft.com/en-us/library/office/dn626019(v=exchg.150).aspx#Anchor_0) pour les boîtes aux lettres Exchange des utilisateurs.

Actuellement, les compléments pour les clients Office suivants sont pris en charge : 

|**Application Office**|**Office 2016 pour Windows**|**Office Online**|**Office 2016 pour Mac**|
|:---------------------|:--------------------------|:--------------|:------------------|
|Word|X|X|X|
|Excel|X|X|X|
|PowerPoint|X|X|X|
|Outlook|Bientôt disponible|Bientôt disponible|Bientôt disponible|

Le centre d’administration ne prend pas en charge les éléments suivants :

- Les compléments qui ciblent Word, Excel, PowerPoint ou Outlook dans Office 2013.
- Le service d’annuaire local.
- Le déploiement de compléments SharePoint.
- Le déploiement de compléments vers Office Online Server.
- Le déploiement de compléments COM/VSTO.

Pour déployer des compléments SharePoint ou des compléments qui ciblent Office 2013, utilisez un [catalogue de compléments SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

>**Important :** les catalogues de compléments SharePoint ne prennent pas en charge les fonctionnalités de complément qui sont implémentées dans le nœud [VersionOverrides](../../reference/manifest/versionoverrides.md) du manifeste de complément, comme les [commandes de complément](../design/add-in-commands.md). 

Pour déployer des compléments COM/VSTO, utilisez ClickOnce ou Windows Installer. Pour plus d’informations, consultez l’article [Déploiement d’une solution Office](https://msdn.microsoft.com/en-us/library/bb386179.aspx).

<!-- Need URL on SOC site.
For more information about requirements, see [centralized deployment eligibility]().
-->

## <a name="publish-an-add-in-via-centralized-deployment"></a>Publication d’un complément via le déploiement centralisé

Pour publier un complément via le déploiement centralisé, procédez comme suit :

1.    Vérifiez que votre organisation répond aux [conditions préalables au déploiement centralisé](#prerequisites-for-centralized-deployment).
2.    Dans la page du centre d’administration Office 365, choisissez **paramètres** > **Services et compléments**.
3.    Sélectionnez **Ajouter un complément Office** en haut de la page. Vous avez le choix parmi les options suivantes :

    - Ajouter un complément à partir de l’Office Store.
    - Sélectionner l’option **Parcourir** pour rechercher votre fichier manifeste (.xml).
    - Entrez l’URL de votre manifeste dans le champ indiqué.

5.    Cliquez sur **Suivant**.
6.    Si vous ajoutez un complément à partir de l’Office Store, sélectionnez le complément. Le complément est désormais activé. 
7.    Sélectionnez **Modifier** pour attribuer le complément à des utilisateurs. 
8.    Recherchez les utilisateurs ou les groupes vers lesquels vous souhaitez déployer le complément, puis cliquez sur **Ajouter** en regard de leur nom.
9.    Cliquez sur **Enregistrer**, passez en revue les paramètres du complément, puis cliquez sur **fermer**.


Si le complément prend en charge les commandes de complément, celles-ci apparaissent dans le ruban de l’application Office pour tous les utilisateurs vers lesquels le complément est déployé. 

Si le complément ne prend pas en charge les commandes de complément, les utilisateurs peuvent l’ajouter à l’aide du bouton **Mes compléments** en procédant comme suit :

1.    Dans Word 2016, Excel 2016 ou PowerPoint 2016, sélectionnez **Insérer** > **Mes compléments**.
2.    Sélectionnez l’onglet **Géré par l’administrateur** dans le fenêtre du complément.
3.    Choisissez le complément, puis cliquez sur **Ajouter**. 

