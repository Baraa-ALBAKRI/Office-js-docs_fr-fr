
# <a name="deploy-and-publish-your-office-add-in"></a>Déploiement et publication de votre complément Office

Vous pouvez utiliser l’une des méthodes pour déployer votre complément Office à des fins de test ou de distribution auprès des utilisateurs.

|**Méthode**|**Use...**|
|:---------|:------------|
|[Chargement de version test](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|Dans le cadre de votre processus de développement, pour tester l’exécution de votre complément sur Windows, Office Online, iPad ou Mac.|
|[Aperçu du centre d’administration Office 365](#office-365-admin-center-preview)|Dans un déploiement cloud ou hybride, pour distribuer votre complément à des utilisateurs de votre organisation.|
|[Office Store]|Pour distribuer publiquement votre complément auprès des utilisateurs.|
|[Catalogue SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|Dans un environnement local, pour distribuer votre complément auprès des utilisateurs de votre organisation.|
|[Serveur Exchange](#outlook-add-in-deployment)|Dans un environnement local ou en ligne, pour distribuer des compléments Outlook à des utilisateurs.|

Les options disponibles dépendent de l’hôte Office que vous ciblez et du type de complément.

>**Remarque :** si vous envisagez de publier votre complément sur l’Office Store, assurez-vous que vous respectez les [stratégies de validation de l’Office Store](https://msdn.microsoft.com/en-us/library/jj220035.aspx). Par exemple, pour obtenir la validation, votre complément doit fonctionner sur toutes les plateformes qui prennent en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](https://dev.office.com/officestore/docs/validation-policies#4-apps-and-add-ins-behave-predictably) et la [page relative à la disponibilité des compléments Office sur les plateformes et les hôtes](https://dev.office.com/add-in-availability)).

## <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Options de déploiement pour les compléments Word, Excel et PowerPoint

| Point d’extension            | Chargement de version test | Aperçu du centre d’administration Office 365 |Office Store| Catalogue SharePoint*  |
|:----------------|:-----------:|:------------------:|:-------------------------------:|:------------:|
| Contenu         | X           | X                  | X                               | X|
| Volet Office       | X           | X                  | X                               | X|
| Commande           | X           | X                  | X                               |  |

&#42; Les catalogues SharePoint ne prennent pas en charge Office 2016 pour Mac.

## <a name="deployment-options-for-outlook-add-ins"></a>Options de déploiement pour les compléments Outlook

| Point d’extension     | Chargement de version test | Serveur Exchange | Office Store |
|:---------|:-----------:|:---------------:|:------------:|
| Application de messagerie | X           | X               | X            |
| Commande  | X           | X               | X            |


Pour plus d’informations sur l’acquisition, l’insertion et l’exécution des compléments par les utilisateurs finals, voir [Commencer à utiliser votre complément Office](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).

## <a name="office-365-admin-center-preview-deployment"></a>Déploiement de l’aperçu du centre d’administration Office 365

Le centre d’administration Office 365 permet aux administrateurs de déployer facilement des compléments Word, Excel et PowerPoint auprès d’utilisateurs ou de groupes au sein de leur organisation. Les compléments déployés via le centre d’administration sont disponibles pour les utilisateurs directement dans leurs applications Office, sans qu’aucune configuration client ne soit requise. Vous pouvez déployer des compléments internes, ainsi que des compléments fournis par des éditeurs de logiciels indépendants via le centre d’administration.

Le centre d’administration prend actuellement en charge les scénarios suivants :

- Déploiement centralisé de nouveaux compléments et de ceux mis à jour pour certaines personnes, des groupes ou une organisation.
- Prise en charge de plusieurs plateformes, y compris Windows et Office Online (Mac bientôt disponible).
- Déploiement en anglais et clients dans le monde entier.
- Déploiement de compléments hébergés sur le cloud.
- Installation automatique de l’application Office au lancement.
- URL des compléments hébergées au sein d’une zone protégée par un pare-feu.
- Déploiement de compléments Office Store (bientôt disponible).

<!--
The admin center also includes a pre-deployment validation checking service.
-->

Les investissements futurs dans des scénarios de déploiement de compléments porteront sur le centre d’administration Office 365. Nous vous recommandons d’utiliser le centre d’administration pour déployer des compléments dans votre organisation, si votre organisation remplit les conditions préalables.

### <a name="prerequisites-for-admin-center-deployment"></a>Conditions préalables pour le déploiement via le centre d’administration 

Vous pouvez déployer des compléments via le centre d’administration si votre organisation répond aux critères suivants :

- Les utilisateurs exécutent Office 2016 build 7070 ou une version ultérieure.
- Les utilisateurs se connectent à Office 2016 avec leur compte professionnel ou scolaire.
- Votre organisation utilise le service d’identité Azure Active Directory (Azure AD).

Le centre d’administration ne prend pas en charge les éléments suivants :

- Les compléments qui ciblent Word, Excel ou PowerPoint dans Office 2013.
- Le service d’annuaires local.
- Le déploiement de compléments SharePoint.
- Le déploiement de compléments vers Office Online Server.
- Le déploiement de compléments COM/VSTO.

Pour déployer des compléments SharePoint ou des compléments qui ciblent Office 2013, utilisez un [catalogue de compléments SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

>**Important :** les catalogues de compléments SharePoint ne prennent pas en charge les fonctionnalités de complément qui sont implémentées dans le nœud [VersionOverrides](../../reference/manifest/versionoverrides.md) du manifeste de complément, comme les [commandes de complément](../design/add-in-commands.md). 

Pour déployer des compléments COM/VSTO, utilisez ClickOnce ou Windows Installer. Pour plus d’informations, voir [Déploiement d’une solution Office](https://msdn.microsoft.com/en-us/library/bb386179.aspx).

## <a name="sharepoint-catalog-deployment"></a>Déploiement d’un catalogue SharePoint

Un catalogue de compléments SharePoint est une collection de sites spéciale que vous pouvez créer pour héberger des compléments Word, Excel et PowerPoint. Les catalogues SharePoint ne prennent pas en charge les nouvelles fonctionnalités de complément mises en œuvre dans le nœud VersionOverrides du manifeste, y compris les commandes de complément. Nous vous recommandons d’utiliser un déploiement centralisé via l’aperçu du centre d’administration si possible. Par défaut, les commandes de complément déployées via un catalogue SharePoint s’ouvrent dans un volet des tâches.

Si vous déployez des compléments dans un environnement local, utilisez un catalogue SharePoint. Pour obtenir des détails, voir l’article sur la [publication de compléments du volet des tâches et de contenu dans un catalogue SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

> **Remarque :** les catalogues SharePoint ne prennent pas en charge Office 2016 pour Mac. Pour déployer des compléments Office sur les clients Mac, vous devez les envoyer à l’[Office Store]. 

## <a name="outlook-add-in-deployment"></a>Déploiement de compléments Outlook

Pour des environnements en ligne et locaux qui n’utilisent pas le service d’identité Azure AD, vous pouvez déployer des compléments Outlook via le serveur Exchange. 

Le déploiement de compléments Outlook nécessite :

- Office 365, Exchange Online ou Exchange Server 2013 ou version ultérieure
- Outlook 2013 ou une version ultérieure

Pour affecter des compléments à des clients, utilisez le centre d’administration Exchange pour télécharger un manifeste directement, à partir d’un fichier ou d’une URL, ou ajoutez un complément à partir de l’Office Store. Pour affecter des compléments à des utilisateurs individuels, vous devez utiliser Exchange PowerShell. Pour plus d’informations, voir [Installation ou suppression de compléments Outlook pour votre organisation](https://technet.microsoft.com/en-us/library/jj943752(v=exchg.150).aspx) sur TechNet.


## <a name="additional-resources"></a>Ressources supplémentaires

- [Déployer et installer des compléments Outlook à des fins de test](../outlook/testing-and-tips.md) 
- [Soumission de compléments et d’applications web dans l’Office Store][Office Store]
- [Instructions de conception pour les compléments Office](../design/add-in-design)
- [Création de compléments efficaces pour l’Office Store](https://msdn.microsoft.com/en-us/library/jj635874.aspx)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/testing-and-troubleshooting.md)

[Office Store]: http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx
[Office Add-in host and platform availability]: http://dev.office.com/add-in-availability
