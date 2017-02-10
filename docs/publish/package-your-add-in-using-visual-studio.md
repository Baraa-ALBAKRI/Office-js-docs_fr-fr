
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>Créer le package de votre complément à l’aide de Visual Studio pour préparer la publication

Votre package de complément Office contient un fichier XML que vous allez utiliser pour publier le complément. Vous devez publier les fichiers de l’application web de votre projet séparément.


## <a name="deploy-your-web-project-and-package-your-add-in-by-using-visual-studio-2015"></a>Déploiement de votre projet web et empaquetage de votre complément à l’aide de Visual Studio 2015



### <a name="to-deploy-your-web-project"></a>Pour déployer votre projet Web


1. Dans l’ **Explorateur de solutions**, ouvrez le menu contextuel du projet d’complément, puis sélectionnez  **Publier**.
    
    La page **Publier votre complément** s’ouvre.
    
2. Dans la liste déroulante **Profil actuel**, sélectionnez un profil ou choisissez **Nouveau …** pour créer un profil.
    
     >**Remarque**  Un profil de publication indique le serveur sur lequel vous effectuez le déploiement, les informations d’identification nécessaires pour se connecter au serveur, les bases de données à déployer, ainsi que d’autres options de déploiement.

    Si vous choisissez  **Nouveau...**, l’Assistant **Créer un profil de publication** s’ouvre. Vous pouvez utiliser cet Assistant pour importer un profil de publication à partir d’un site web d’hébergement comme Microsoft Azure ou créer un profil et ajouter votre serveur, vos informations d’identification et d’autres paramètres, comme décrit dans la procédure suivante.
    
    Pour plus d’informations sur l’importation et la création de profils de publication, voir [Création d’un profil de publication](http://msdn.microsoft.com/en-us/library/dd465337.aspx#creating_a_profile).
    
3. Sur la page  **Publier votre complément**, cliquez sur le lien  **Déployer votre projet Web**.
    
    The  **Publish Web** dialog box appears. For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](http://msdn.microsoft.com/en-us/library/dd465337.aspx).
    

### <a name="to-package-your-add-in"></a>Empaquetage de votre complément


1. Sur la page  **Publier votre complément**, cliquez sur le lien  **Empaqueter le complément**.
    
    L’Assistant **Publication des compléments SharePoint et Office** apparaît.
    
2. Dans la liste déroulante  **Où votre site web est-il hébergé ?**, sélectionnez ou saisissez l’URL du site web qui hébergera les fichiers de contenu de votre complément, puis cliquez sur  **Terminer**.
    
    You have to specify an address that begins with the HTTPS prefix to complete this wizard. In general, using an HTTPS endpoint for your website is the best approach, but it is not required if you don't plan to publish your add-in to the Office Store. After the package is created, you can open the manifest in Notepad and replace the HTTPS prefix of your website with an HTTP prefix. For more information, see [Why do my add-ins have to be SSL-secured?](http://msdn.microsoft.com/en-us/library/jj591603#bk_q7). 
    
     >**Remarque**  Les sites web Azure fournissent automatiquement un point de terminaison HTTPS.

    Visual Studio génère les fichiers nécessaires à la publication de votre complément, puis ouvre le dossier de sortie de publication. 
    
Si vous prévoyez de soumettre votre complément à l’Office Store, vous pouvez cliquer sur le lien **Effectuer un test de validation** pour identifier les problèmes susceptibles d’empêcher votre complément d’être accepté. Vous devez régler tous ces problèmes avant de soumettre votre complément au magasin.

Vous pouvez désormais télécharger votre manifeste XML à l’emplacement approprié pour [publier votre complément](../publish/publish.md). Le manifeste XML se trouve dans  `OfficeAppManifests` dans le dossier `app.publish`. Par exemple :

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="additional-resources"></a>Ressources supplémentaires



- [Publier votre complément Office](../publish/publish.md)
    
- [Soumission des compléments SharePoint et Office, ainsi que des applications web Office 365 dans l’Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    
